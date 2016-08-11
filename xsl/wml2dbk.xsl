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
  xmlns="http://docbook.org/ns/docbook"
  version="2.0" 
  exclude-result-prefixes = "w xs dbk r rel tr m mc xlink docx2hub wp">

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

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- named Templates -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template name="handle-field-function">
    <xsl:param name="nodes" as="element(*)*"/>
    <xsl:param name="is-multi-para" as="xs:boolean" select="false()"/>
    <xsl:variable name="field-code" select="normalize-space(string-join($nodes/w:instrText, ''))" as="xs:string"/>
    <xsl:choose>
      <!-- construct link element -->
      <xsl:when test="matches($field-code, '^REF\s([_A-Za-z]+\d+)\s?.+\\h(\s|$)')"><!-- hyperlink to bookmark -->
        <xsl:variable name="linkend" select="replace($field-code, '^REF\s([_A-Za-z]+\d+)\s?.+$', '$1')" as="xs:string"/>
        <link linkend="{$linkend}">
          <xsl:apply-templates select="$nodes" mode="#current"/>
        </link>    
      </xsl:when>
      <xsl:when test="matches($field-code, '^REF\s.+')"><!-- other refs, e.g., to variable that was set using SET.
        We assume that the value of SET is identical with the content, so dissolve this. Is this assumption justified? --> 
        <xsl:apply-templates select="$nodes" mode="#current"/>
      </xsl:when>
      <xsl:when test="$is-multi-para">
        <xsl:choose>
          <xsl:when test="$nodes[1][self::w:p[count(w:r/w:fldChar[@w:fldCharType = 'begin']) = 2]]">
            <xsl:if test="not($nodes[last()][self::w:p[count(w:r/w:fldChar) = 1 and w:r/w:fldChar[@w:fldCharType = 'end']]])">
              <xsl:call-template name="signal-error" xmlns="">
                <xsl:with-param name="error-code" select="'W2D_010'"/>
                <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
                <xsl:with-param name="hash">
                  <value key="xpath"><xsl:value-of select="$nodes[last()]/@srcpath"/></value>
                  <value key="level">INT</value>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:if>
            <xsl:variable name="split" select="$nodes[w:r/w:fldChar[@w:fldCharType = 'begin'] and w:r/w:fldChar[@w:fldCharType = 'end']
                                               and count(w:r/w:fldChar[@w:fldCharType = 'end']) = count(w:r/w:fldChar[@w:fldCharType = 'begin'])
                                               and w:r[w:fldChar/@w:fldCharType = 'begin'][1]/preceding-sibling::w:r[w:fldChar/@w:fldCharType = 'end']
                                               and not(.//w:t)]"/>
            <xsl:apply-templates select="$nodes[not(position() = last())] except $split" mode="wml-to-dbk"/>
          </xsl:when>
          <xsl:when test="$nodes[1][self::w:p]">
            <xsl:variable name="first-node">
              <xsl:element name="{$nodes[1]/name()}">
                <xsl:copy-of select="$nodes[1]/@*"/>
                <xsl:copy-of select="$nodes[1]/node()"/>
                <w:r>
                  <w:fldChar w:fldCharType="end"/>
                </w:r>
              </xsl:element>
            </xsl:variable>
            <xsl:apply-templates select="$first-node" mode="wml-to-dbk"/>
            <xsl:apply-templates select="$nodes[position() &gt; 1 and not(position() = last())]" mode="wml-to-dbk"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="$nodes" mode="wml-to-dbk"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="count($nodes/w:fldChar[@w:fldCharType = 'begin']) - count($nodes/w:fldChar[@w:fldCharType = 'end']) gt 0">
          <xsl:call-template name="signal-error" xmlns="">
            <xsl:with-param name="error-code" select="'W2D_011'"/>
            <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
            <xsl:with-param name="hash">
              <value key="xpath"><xsl:value-of select="$nodes[1]/@srcpath"/></value>
              <value key="level">INT</value>
              <value key="info-text"><xsl:value-of select="$nodes//text()"/></value>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
        <xsl:choose>
          <xsl:when test="count($nodes/w:fldChar[@w:fldCharType = 'begin']) eq 1">
            <xsl:choose>
              <xsl:when test="$nodes[1][self::w:r[w:fldChar[@w:fldCharType = 'begin']]]">
                <xsl:choose>
                  <xsl:when test="not($nodes[w:fldChar[@w:fldCharType = 'separate']])">
                    <xsl:variable name="end" select="$nodes[w:fldChar[@w:fldCharType = 'end']]"/>
                    <xsl:apply-templates select="($nodes[position() &gt; 1 and . &lt;&lt; $end])[1]" mode="wml-to-dbk">
                      <xsl:with-param name="instrText" select="string-join($nodes[position() &gt; 1 and . &lt;&lt; $end]//text(), '')" tunnel="yes"/>
                      <xsl:with-param name="nodes" select="$nodes[position() &gt; 1 and . &lt;&lt; $end]" tunnel="yes"/>
                      <xsl:with-param name="text" select="()" tunnel="yes"/>
                    </xsl:apply-templates>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:variable name="sep" select="$nodes[w:fldChar[@w:fldCharType = 'separate']]"/>
                    <xsl:variable name="end" select="$nodes[w:fldChar[@w:fldCharType = 'end']]"/>
                    <xsl:apply-templates select="($nodes[position() &gt; 1 and . &lt;&lt; $sep])[1]" mode="wml-to-dbk">
                      <xsl:with-param name="instrText" select="string-join($nodes[position() &gt; 1 and . &lt;&lt; $sep]//text(), '')" tunnel="yes"/>
                      <xsl:with-param name="nodes" select="$nodes[position() &gt; 1 and . &lt;&lt; $sep]" tunnel="yes"/>
                      <xsl:with-param name="text" select="$nodes[. &gt;&gt; $sep and . &lt;&lt; $end]" tunnel="yes"/>
                    </xsl:apply-templates>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="signal-error" xmlns="">
                  <xsl:with-param name="error-code" select="'W2D_012'"/>
                  <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
                  <xsl:with-param name="hash">
                    <value key="xpath"><xsl:value-of select="$nodes[1]/@srcpath"/></value>
                    <value key="level">INT</value>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="count($nodes/w:fldChar[@w:fldCharType = 'begin']) gt 1">
            <xsl:choose>
              <xsl:when test="$nodes/w:fldChar[@w:fldCharType = 'begin'][following::w:fldChar[@w:fldCharType = ('begin','end')][1][self::w:fldChar[@w:fldCharType = 'begin']]]">
                <xsl:call-template name="handle-nested-field-functions">
                  <xsl:with-param name="nodes" select="$nodes"/>
                  <xsl:with-param name="depth" select="0"/>
                </xsl:call-template>
                </xsl:when>
            <xsl:otherwise>
                <xsl:for-each-group select="$nodes" group-starting-with="w:r[w:fldChar[@w:fldCharType = 'begin']]">
                  <xsl:choose>
                    <xsl:when test="current-group()[1][self::w:r[w:fldChar[@w:fldCharType = 'begin']]]">
                      <xsl:apply-templates select="(current-group()[w:instrText])[1]" mode="wml-to-dbk">
                        <xsl:with-param name="instrText" select="string-join(current-group()//text()[parent::w:instrText], '')" tunnel="yes" as="xs:string?"/>
                        <xsl:with-param name="nodes" select="current-group()[descendant::w:instrText]" tunnel="yes" as="element(*)*"/>
                        <xsl:with-param name="text" select="current-group()[.//text()[parent::w:t] or .//w:tab or .//w:br or .//w:pict]" tunnel="yes" as="element(*)*"/>
                      </xsl:apply-templates>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:call-template name="signal-error" xmlns="">
                        <xsl:with-param name="error-code" select="'W2D_012'"/>
                        <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
                        <xsl:with-param name="hash">
                          <value key="xpath"><xsl:value-of select="current-group()[1]/@srcpath"/></value>
                          <value key="level">INT</value>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each-group>    
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_013'"/>
              <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="$nodes[1]/@srcpath"/></value>
                <value key="level">INT</value>
                <value key="info-text"><xsl:value-of select="count($nodes/w:fldChar[@w:fldCharType = 'begin'])"/></value>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="handle-nested-field-functions">
    <xsl:param name="nodes" as="node()*"/>
    <xsl:param name="depth" as="xs:integer"/>
    <xsl:choose>
      <xsl:when test="$depth lt 16">
        <xsl:choose>
          <xsl:when test="not($nodes/w:fldChar[@w:fldCharType='begin']) and $nodes//w:instrText and matches(string-join($nodes//w:instrText//text(),''),'^[A-Z\.]*[0-9]*$')">
        <xsl:copy-of select="$nodes//w:instrText/text()"/>
      </xsl:when>
      <xsl:when test="not($nodes/w:fldChar[@w:fldCharType='begin']) and (every $i in $nodes satisfies $i[self::w:r[w:instrText[text()]]])">
            <xsl:copy-of select="$nodes//w:instrText/text()"/>
          </xsl:when>
          <xsl:when test="not($nodes/w:fldChar[@w:fldCharType='begin'])">
        <xsl:copy-of select="$nodes"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="new-nodes" as="document-node(element(dbk:temp))">
          <xsl:document><temp>
            <xsl:for-each-group select="$nodes" group-starting-with="w:r[w:fldChar[@w:fldCharType='begin'][following::w:fldChar[@w:fldCharType = ('begin','end')][1][self::w:fldChar[@w:fldCharType = 'end']]]]">
              <xsl:for-each-group select="current-group()" group-ending-with="w:r[w:fldChar[@w:fldCharType='end']]">
                <xsl:choose>
                  <xsl:when test="current-group()[1][self::w:r[
                                                       w:fldChar[
                                                         @w:fldCharType='begin'
                                                       ][
                                                         following::w:fldChar[
                                                           @w:fldCharType = ('begin','end')
                                                         ][1][
                                                           self::w:fldChar[@w:fldCharType = 'end']
                                                         ]
                                                       ]
                                                     ]
                                                   ] 
                                  and 
                                  current-group()[last()][
                                                    self::w:r[w:fldChar[@w:fldCharType='end']]
                                                 ]">
                    <xsl:variable name="prelim" as="node()*">
                      <xsl:apply-templates select="(current-group()[w:instrText])[1]" mode="wml-to-dbk">
                        <xsl:with-param name="instrText" select="string-join(current-group()//text()[parent::w:instrText], '')" tunnel="yes" as="xs:string?"/>
                        <xsl:with-param name="nodes" select="current-group()[descendant::w:instrText]" tunnel="yes" as="element(*)*"/>
                        <xsl:with-param name="text" select="current-group()[.//text()[parent::w:t] or .//w:tab or .//w:br or .//w:pict or descendant-or-self::dbk:*]" tunnel="yes" as="element(*)*"/>
                      </xsl:apply-templates>
                    </xsl:variable>
                    <xsl:for-each select="$prelim">
                      <xsl:choose>
                        <xsl:when test="self::text()">
                          <w:r>
                            <!-- Not sure whether we can safely surround text output (SYMBOL field function processing output)
                              with instrText. If we don't, SYMBOL within XE will be discarded -->
                            <w:instrText>
                              <xsl:sequence select="."/>
                            </w:instrText>
                          </w:r>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:sequence select="."/>
                        </xsl:otherwise>
                      </xsl:choose>  
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:copy-of select="current-group()"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each-group>
            </xsl:for-each-group>
          </temp>
          </xsl:document>
        </xsl:variable>
        <xsl:call-template name="handle-nested-field-functions">
          <xsl:with-param name="nodes" select="$new-nodes/*/node()"/>
        <xsl:with-param name="depth" select="$depth + 1"/>
            </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="signal-error" xmlns="">
            <xsl:with-param name="error-code" select="'W2D_094'"/>
            <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
            <xsl:with-param name="hash">
              <value key="xpath"><xsl:value-of select="$nodes[1]/@srcpath"/></value>
              <value key="level">ERR</value>
              <value key="info-text"><xsl:value-of select="$nodes//text()"/></value>
              <value key="pi">W2D_094 <xsl:value-of select="$nodes//text()"/></value>
            </xsl:with-param>
          </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="check-field-functions">
    <xsl:param name="nodes" as="element(*)*"/>
    <xsl:for-each-group select="*"
      group-adjacent="count(preceding::w:fldChar[@w:fldCharType = 'begin'])
                      - count(preceding::w:fldChar[@w:fldCharType = 'end'])
                      + (if (w:r/w:fldChar[@w:fldCharType = 'begin'])
                      then (count(w:r/w:fldChar[@w:fldCharType = 'begin']) - count(w:r/w:fldChar[@w:fldCharType = 'end']))
                      else 0)">
      <xsl:choose>
        <xsl:when test="current-grouping-key() &gt; 0">
          <xsl:for-each-group select="current-group()" group-starting-with="*[count(preceding::w:fldChar[@w:fldCharType = 'begin'])
                                                                            - count(preceding::w:fldChar[@w:fldCharType = 'end']) = 0]">
            <xsl:call-template name="handle-field-function">
              <xsl:with-param name="nodes" select="current-group()"/>
              <xsl:with-param name="is-multi-para" select="true()"/>
            </xsl:call-template>
            <xsl:if test="current-group()[last()]/w:r[w:fldChar[@w:fldCharType = 'end']][last()]/following-sibling::*">
              <!-- verlorengegangenen Knoten ohne @w:fldChar reproduzieren -->
              <xsl:variable name="saved-last-node">
                <xsl:apply-templates select="current-group()[position() = last()]" mode="rescue-node"/>
              </xsl:variable>
              <xsl:apply-templates select="$saved-last-node" mode="wml-to-dbk"/>
            </xsl:if>
          </xsl:for-each-group>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="current-group()" mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each-group>
  </xsl:template>

  <xsl:template match="w:r[w:fldChar[@w:fldCharType = 'end'] and not(following-sibling::w:r[w:fldChar])]" mode="rescue-node">
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- Mode: pre-process -->
  <!-- ================================================================================ -->

  <!-- Ende von Feldfunktionen ueber mehrere Absaetze in einzelnen Absatz packen -->
  <!-- Grund: wenn in dem gleichen Absatz eine neue Feldfunktion beginnt, liefert check-field-functions falsche Gruppen -->
  <xsl:template match="w:p[
                         w:r[w:fldChar][1][count(w:fldChar) = 1]
                         /w:fldChar[@w:fldCharType='end']
                       ][count(w:r[w:fldChar]) gt 1]" mode="docx2hub:separate-field-functions">
    <xsl:variable name="attribute-names" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="name(.)"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="attribute-values" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="."/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="pPr" as="node() *">
      <xsl:apply-templates select="w:pPr" mode="#current"/>
    </xsl:variable>
    <xsl:for-each-group select="*" group-ending-with="w:r[w:fldChar][1]">
      <w:p>
        <xsl:for-each select="$attribute-names">
          <xsl:variable name="pos" select="position()"/>
          <xsl:attribute name="{.}" select="$attribute-values[position() eq $pos]"/>
        </xsl:for-each>
        <xsl:copy-of select="$pPr"/>
        <xsl:apply-templates select="current-group()[not(self::w:pPr)]" mode="#current"/>
      </w:p>
    </xsl:for-each-group>
  </xsl:template>

  <xsl:template match="w:p[
                         w:r[last()][w:fldChar][count(w:fldChar) = 1]
                         /w:fldChar[@w:fldCharType='end']
                       ][
                         count(w:r[w:fldChar[@w:fldCharType='end']])
                         gt
                         count(w:r[w:fldChar[@w:fldCharType='begin']])
                       ][
                         count(w:r[w:fldChar[@w:fldCharType='end']]) gt 1
                       ]" mode="docx2hub:separate-field-functions">
    <xsl:variable name="attribute-names" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="name(.)"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="attribute-values" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="."/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="pPr" as="node() *">
      <xsl:apply-templates select="w:pPr" mode="#current"/>
    </xsl:variable>
    <xsl:for-each-group select="*" group-ending-with="w:r[position() lt last()][w:fldChar][last()]">
      <w:p>
        <xsl:for-each select="$attribute-names">
          <xsl:variable name="pos" select="position()"/>
          <xsl:attribute name="{.}" select="$attribute-values[position() eq $pos]"/>
        </xsl:for-each>
        <xsl:copy-of select="$pPr"/>
        <xsl:apply-templates select="current-group()[not(self::w:pPr)]"/>
      </w:p>
    </xsl:for-each-group>
  </xsl:template>
  
  <!-- Links that stretch across para boundaries: Insert new instrTexts so that there are individual links per para.
  Might need to put it into a mode of its own. -->
  <!-- this is work in progress
  <xsl:template match="/" mode="docx2hub:separate-field-functions">
    <xsl:variable name="link-begins" as="element(w:fldChar)*" 
      select="for $it in .//w:instrText[matches(., '\s*REF\s+')]
              return $it/preceding::w:fldChar[1]"/>
    <xsl:variable name="link-ends" as="element(w:fldChar)*" 
      select=".//w:fldChar[@w:fldCharType = 'end']
                          [exists(key('docx2hub:item-by-id', @linkend) intersect $link-begins)]"/>
    <xsl:next-match>
      <xsl:with-param name="link-begins" select="$link-begins" tunnel="yes"/>
      <xsl:with-param name="link-ends" select="$link-ends" tunnel="yes"/>
    </xsl:next-match>
  </xsl:template>
  
  <xsl:template match="w:p" priority="2" mode="docx2hub:separate-field-functions">
    <xsl:param name="link-begins" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="link-ends" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:variable name="last-contained-begin" select="($link-begins intersect .//w:fldChar)[last()]" as="element(w:fldChar)?"/>
    <xsl:variable name="last-contained-corresponding-end" as="element(w:fldChar)?"
      select="key('docx2hub:linking-item-by-id', $last-contained-begin/@xml:id) intersect $link-ends"/>
    <xsl:variable name="last-outside-begin" select="($link-begins[. &lt;&lt; current()])[last()]" as="element(w:fldChar)?"/>
    <xsl:variable name="last-outside-corresponding-end" as="element(w:fldChar)?"
      select="key('docx2hub:linking-item-by-id', $last-outside-begin/@xml:id) intersect $link-ends"/>
    <xsl:message select="$last-outside-begin, exists($last-outside-corresponding-end intersect .//w:fldChar)"></xsl:message>
    <xsl:choose>
      <xsl:when test="exists($last-outside-begin) and exists($last-outside-corresponding-end intersect .//w:fldChar)">
        <xsl:comment>hurz</xsl:comment>
        <xsl:copy>
          <xsl:apply-templates select="@*" mode="#current"/>
          <xsl:apply-templates select="$last-contained-begin, node()" mode="#current">
            <xsl:with-param name="id-suffix" select="generate-id()" tunnel="yes"/>
            <xsl:with-param name="link-begin" select="$last-contained-begin" tunnel="yes"/>
            <xsl:with-param name="link-end" select="$last-outside-corresponding-end" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:copy>
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:fldChar[@w:fldCharType = 'begin']" mode="docx2hub:separate-field-functions">
    <xsl:param name="id-suffix" as="xs:string?" tunnel="yes"/>
    <xsl:param name="link-begin" as="element(w:fldChar)?" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test="$link-begin is current()">
        <xsl:copy>
          <xsl:attribute name="xml:id" select="concat(@xml:id, '_', $id-suffix)"/>
          <xsl:apply-templates select="@* except @xml:id"/>
        </xsl:copy>
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:fldChar[@w:fldCharType = 'end']" mode="docx2hub:separate-field-functions">
    <xsl:param name="id-suffix" as="xs:string?" tunnel="yes"/>
    <xsl:param name="link-end" as="element(w:fldChar)?" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test="$link-end is current()">
        <xsl:copy>
          <xsl:attribute name="xml:id" select="concat(@xml:id, '_', $id-suffix)"/>
          <xsl:apply-templates select="@* except @xml:id"/>
        </xsl:copy>
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

-->

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
  <xsl:template match="/dbk:*" mode="docx2hub:separate-field-functions">
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
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:call-template name="check-field-functions">
        <xsl:with-param name="nodes" select="*"/>
      </xsl:call-template>
    </xsl:copy>
  </xsl:template>
  
  <!-- paragraphs (w:p) -->

  <xsl:variable name="docx2hub:allowed-para-element-names" as="xs:string+"
    select="('w:r', 'w:pPr', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:smartTag', 'w:commentRangeStart', 'w:commentRangeEnd', 'w:proofErr', 'w:hyperlink', 'w:del', 'w:ins', 'w:fldSimple', 'm:oMathPara', 'm:oMath')" />

  <xsl:template match="w:p" mode="wml-to-dbk">
    <xsl:element name="para">
      <xsl:apply-templates select="@* except @*[matches(name(),'^w:rsid')]" mode="#current"/>
<!--      <xsl:if test=".//w:r">-->
        <xsl:sequence select="tr:insert-numbering(.)"/>
      <!--</xsl:if>-->
      <!-- Only necessary in tables? They'll get lost otherwise. -->
      <xsl:variable name="bookmarkstart-before-p" as="element(w:bookmarkStart)*"
        select="preceding-sibling::w:bookmarkStart[. &gt;&gt; current()/preceding-sibling::*[not(self::w:bookmarkStart)][1]]"/>
      <xsl:variable name="bookmarkstart-before-tc" as="element(w:bookmarkStart)*"
        select="parent::w:tc[current() is w:p[1]]/preceding-sibling::w:bookmarkStart[. &gt;&gt; current()/parent::w:tc/preceding-sibling::*[not(self::w:bookmarkStart)][1]]"/>
      <xsl:variable name="bookmarkstart-before-tr" as="element(w:bookmarkStart)*"
        select="parent::w:tc/parent::w:tr[current() is (w:tc/w:p)[1]]/preceding-sibling::w:bookmarkStart[. &gt;&gt; current()/parent::w:tc/parent::w:tr/preceding-sibling::*[not(self::w:bookmarkStart)][1]]"/>
      <xsl:variable name="bookmarkend-after-p" as="element(w:bookmarkEnd)*"
        select="following-sibling::w:bookmarkEnd[. &lt;&lt; current()/following-sibling::*[not(self::w:bookmarkEnd)][1]]"/>
      <xsl:variable name="bookmarkend-after-tc" as="element(w:bookmarkEnd)*"
        select="parent::w:tc[current() is w:p[1]]/following-sibling::w:bookmarkEnd[. &lt;&lt; current()/parent::w:tc/following-sibling::*[not(self::w:bookmarkEnd)][1]]"/>
      <xsl:variable name="bookmarkend-after-tr" as="element(w:bookmarkEnd)*"
        select="parent::w:tc/parent::w:tr[current() is (w:tc/w:p)[1]]/following-sibling::w:bookmarkEnd[. &lt;&lt; current()/parent::w:tc/parent::w:tr/following-sibling::*[not(self::w:bookmarkEnd)][1]]"/>

      <xsl:apply-templates select="$bookmarkstart-before-p | $bookmarkstart-before-tc | $bookmarkstart-before-tr" mode="wml-to-dbk-bookmarkStart"/>
      <xsl:choose>
        <xsl:when test="w:r[w:fldChar]">
          <xsl:call-template name="inline-field-function"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="node() except dbk:tabs" mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:apply-templates select="$bookmarkend-after-p | $bookmarkend-after-tc | $bookmarkend-after-tr" mode="wml-to-dbk-bookmarkEnd"/>
    </xsl:element>
  </xsl:template>

  <xsl:template name="inline-field-function">
    <xsl:variable name="starts" select="count(w:r[w:fldChar/@w:fldCharType = 'begin'])"/>
    <xsl:variable name="ends" select="count(w:r[w:fldChar/@w:fldCharType = 'end'])"/>
    <xsl:variable name="seps" select="count(w:r[w:fldChar/@w:fldCharType = 'separate'])"/>
    <xsl:if test="$starts lt $ends">
      <xsl:call-template name="signal-error" xmlns="">
        <xsl:with-param name="error-code" select="'W2D_014'"/>
        <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
        <xsl:with-param name="hash">
          <value key="xpath"><xsl:value-of select="@srcpath"/></value>
          <value key="level">INT</value>
          <value key="info-text"><xsl:value-of select="."/></value>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
    <xsl:for-each-group select="* except dbk:tabs" 
      group-adjacent="if (count(self::w:r[w:fldChar/@w:fldCharType='begin'])
                      + count(preceding-sibling::w:r[w:fldChar/@w:fldCharType='begin'])
                      (: - count(self::w:r[w:fldChar/@w:fldCharType='end']) :)
                      - count(preceding-sibling::w:r[w:fldChar/@w:fldCharType='end']) &gt; 0) then true() else false()">
      <xsl:choose>
        <xsl:when test="current-grouping-key()">
          <xsl:call-template name="handle-field-function">
            <xsl:with-param name="nodes" select="current-group()"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="current-group()" mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each-group>
  </xsl:template>

  <!-- Verlauf -->
  <xsl:template match="w:del" mode="wml-to-dbk">
    <!-- gelöschten Text wegwerfen -->
  </xsl:template>

  <xsl:template match="w:ins" mode="wml-to-dbk">
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
  
  <xsl:key name="docx2hub:bookmarkStart-by-name" match="w:bookmarkStart[@w:name]" use="docx2hub:normalize-name-for-id(@w:name)"/>
  
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
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:choose>
        <xsl:when test="w:r[w:fldChar]">
          <xsl:call-template name="inline-field-function"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="*" mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
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
    <xsl:variable name="key-name" as="xs:string"
      select="if (ancestor::w:footnote)
              then 'footnote-rel-by-id'
              else if (ancestor::w:comment) 
                then 'comment-rel-by-id'
                else 'doc-rel-by-id'" />
    <xsl:variable name="value" select="."/>
    <xsl:variable name="rel-item" select="key($key-name, current(), $root)" as="element(rel:Relationship)" />
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
          <xsl:copy-of select="key('style-by-name', @role)/(@xml:lang, @css:direction, @docx2hub:rtl-lang)"/>
          <xsl:copy-of select="@xml:lang except $context, ../@css:direction, ../@docx2hub:rtl-lang"/>
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
          <xsl:copy-of select="key('style-by-name', @role)/(@xml:lang, @css:direction, @docx2hub:rtl-lang)"/>
          <xsl:copy-of select="@docx2hub:rtl-lang except $context, ../@css:direction, ../@xml:lang"/>
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

  <!-- instrText (w:instrText)  [not(../preceding-sibling::*[w:instrText])] -->
  <xsl:template match="w:instrText" mode="wml-to-dbk" priority="20">
    <xsl:param name="instrText" as="xs:string?" tunnel="yes"/>
    <xsl:param name="text" as="element(*)*" tunnel="yes"/>
    <xsl:param name="nodes" as="element(*)*" tunnel="yes"/>
    <xsl:variable name="tokens" as="xs:string*">
      <xsl:analyze-string select="($instrText, ' ')[ . ne ''][1]" regex="&quot;(.*?)&quot;">
        <xsl:matching-substring>
          <xsl:sequence select="regex-group(1)"/>
        </xsl:matching-substring>
        <xsl:non-matching-substring>
          <xsl:sequence select="tokenize(., '\s+')[normalize-space(.)]"/>
        </xsl:non-matching-substring>
      </xsl:analyze-string>
    </xsl:variable>
    <!--<xsl:variable name="tokens" select="tokenize(normalize-space($instrText), ' ')" as="xs:string*"/>-->
    <xsl:variable name="func" select="doc('')//tr:field-functions/tr:field-function[@name = $tokens[1]]" as="element(tr:field-function)?"/>
    <xsl:choose>
      <xsl:when test="not($func)">
        <xsl:choose>
          <xsl:when test="$tokens[1] = 'SYMBOL'">
            <!-- Template in sym.xsl -->
            <xsl:call-template name="create-symbol">
              <xsl:with-param name="tokens" select="$tokens"/>
              <xsl:with-param name="context" select=".."/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$tokens[1] = ('XE', 'xe')">
            <xsl:call-template name="handle-index">
              <xsl:with-param name="instr" select="$instrText"/>
              <xsl:with-param name="text" select="$text"/>
              <xsl:with-param name="nodes" select="$nodes"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$tokens[1] = ('EQ','eq','FORMCHECKBOX')">
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_045'"/>
              <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="@srcpath"/></value>
                <value key="level">WRN</value>
                <value key="info-text"><xsl:value-of select="$instrText"/></value>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:apply-templates select="$text" mode="#current"/>
          </xsl:when>
          <xsl:when test="$tokens[1] = 'INCLUDEPICTURE'">
            <xsl:choose>
              <!-- figures are preferably handled by looking at the relationships 
              because INLCUDEPICTURE is more like a history of all locations where
              the image was once included from.  
              Because there may be multiple INCLUDEPICTURES, we ignore them not only
              if the w:pict is contained in a field function, but if there is any 
              w:pict in the current paragraph. Is this assumption justified?
              -->
              <xsl:when test="$nodes/ancestor::w:p//w:pict">
                <xsl:apply-templates select="$text" mode="#current"/>    
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="handle-figures">
                  <xsl:with-param name="instr" select="$instrText"/>
                  <xsl:with-param name="text" select="$text"/>
                  <xsl:with-param name="nodes" select="$nodes"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="$tokens[1] = 'HYPERLINK'">
            <xsl:variable name="without-options" select="$tokens[not(matches(., '\\[lo]'))]" as="xs:string+"/>
            <xsl:variable name="local" as="xs:boolean" select="$tokens = '\l'"/>
            <xsl:variable name="target" select="replace($without-options[2], '(^&quot;|&quot;$)', '')"/>
            <xsl:variable name="tooltip" select="replace($without-options[3], '(^&quot;|&quot;$)', '')"/>
            <link docx2hub:field-function="yes">
              <xsl:attribute name="{if ($local) then 'linkend' else 'xlink:href'}" select="$target"/>
              <xsl:if test="$tooltip">
                <xsl:attribute name="xlink:title" select="$tooltip"/>
              </xsl:if>
              <xsl:apply-templates select="($nodes//@srcpath)[1], $text" mode="#current"/>
            </link>
          </xsl:when>
          <xsl:when test="$tokens[1] = 'SET'">
            <xsl:if test="$field-vars='yes'">
              <keyword role="{concat('fieldVar_',$tokens[2])}" docx2hub:field-function="yes">
                <xsl:value-of select="$tokens[3]"/>    
              </keyword>
              </xsl:if>
          </xsl:when>
          <xsl:when test="matches($instrText,'^[\s&#160;]*$')">
            <xsl:apply-templates select="$text" mode="#current"/>
          </xsl:when>
          <xsl:when test="$tokens[1] = 'PRINT'">
            <xsl:processing-instruction name="PRINT" select="string-join($tokens[position() gt 1], ' ')"/>
          </xsl:when>
          <xsl:when test="$tokens[1] = 'AUTOTEXT'">
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_045'"/>
              <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="@srcpath"/></value>
                <value key="level">WRN</value>
                <value key="info-text"><xsl:value-of select="$instrText"/></value>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:apply-templates select="$text" mode="#current"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_040'"/>
              <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="@srcpath"/></value>
                <value key="level">INT</value>
                <value key="info-text"><xsl:value-of select="$instrText"/></value>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:apply-templates select="$text" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$func/@element">
            <xsl:element name="{$func/@element}">
              <xsl:attribute name="docx2hub:field-function" select="'yes'"/>
              <xsl:if test="$func/@attrib">
                <xsl:attribute name="{$func/@attrib}" select="replace($tokens[position() = $func/@value], '&quot;', '')"/>
                <xsl:if test="$func/@role">
                  <xsl:attribute name="role" select="$func/@role"/>
                </xsl:if>
                <xsl:apply-templates select="$text" mode="#current"/>
              </xsl:if>
            </xsl:element>
          </xsl:when>
          <xsl:when test="$func/@destroy = 'yes'">
            <xsl:if test="$text[descendant::w:fldChar or descendant-or-self::*[@docx2hub:field-function]]">
              <xsl:apply-templates select="$text" mode="#current"/>
            </xsl:if> 
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="$text" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <tr:field-functions>
    <tr:field-function name="INDEX" destroy="yes"/>
    <tr:field-function name="NOTEREF" element="link" attrib="linkend" value="2"/>
    <tr:field-function name="PAGE"/>
    <tr:field-function name="PAGEREF" element="link" attrib="linkend" role="page" value="2"/>
    <tr:field-function name="RD"/>
    <tr:field-function name="REF"/>
    <tr:field-function name="ADVANCE"/>
    <tr:field-function name="QUOTE"/>
    <tr:field-function name="SEQ"/>
    <tr:field-function name="STYLEREF"/>
    <tr:field-function name="USERPROPERTY" destroy="yes"/>
    <tr:field-function name="TOC" destroy="yes"/>
    <tr:field-function name="\IF"/>
  </tr:field-functions>

  <xsl:template name="handle-figures">
    <xsl:param name="inline" select="false()" as="xs:boolean" tunnel="yes"/>
    <xsl:param name="instr" as="xs:string?"/>
    <xsl:param name="text" as="element(*)*"/>
    <xsl:param name="nodes" as="element(*)*"/>
    <xsl:variable name="text-tokens" select="for $x in $nodes//text() return $x"/>
    <xsl:element name="mediaobject">
      <xsl:attribute name="docx2hub:field-function" select="'yes'"/>
      <xsl:apply-templates select="($nodes//@srcpath)[1]" mode="#current"/>
      <imageobject>
        <imagedata fileref="{if (tokenize($instr, ' ')[matches(.,'^&#x22;.*&#x22;$')]) then replace(tokenize($instr, ' ')[matches(.,'^&#x22;.*&#x22;$')][1],'&#x22;','') else if (matches($instr,'&#x22;.*&#x22;')) then tokenize($instr,'&#x22;')[2] else tokenize($instr, ' ')[2]}"/>
      </imageobject>
    </xsl:element>
  </xsl:template>

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

