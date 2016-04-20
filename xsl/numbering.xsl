<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:tr="http://transpect.io"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0"
  exclude-result-prefixes="w xs dbk tr docx2hub">

  <!-- This mode is called from docx2hub:remove-redundant-run-atts as a collateral -->

  <xsl:function name="docx2hub:follow-styleLinks" as="element(w:abstractNum)?">
    <xsl:param name="abstractNum" as="element(w:abstractNum)?"/>
    <xsl:variable name="resolved" as="element(w:abstractNum)?"
      select="key('abstractNum-by-styleLink', $abstractNum/w:numStyleLink/@w:val, $root)"/>
    <xsl:choose>
      <xsl:when test="exists($resolved)">
        <xsl:sequence select="docx2hub:follow-styleLinks($resolved)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$abstractNum"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="w:numId" mode="docx2hub:abstractNum">
    <xsl:param name="ilvl" as="xs:integer"/>
    <xsl:variable name="abstractNum" as="element(w:abstractNum)?"  
      select="docx2hub:follow-styleLinks(
                key(
                  'abstract-numbering-by-id', 
                  key(
                    'numbering-by-id', 
                    @w:val,
                    $root
                  )/w:abstractNumId/@w:val, 
                  $root
                )
              )"/>
    <xsl:variable name="lvl" as="element(w:lvl)?" select="$abstractNum/w:lvl[@w:ilvl = $ilvl]"/>
    <xsl:variable name="lvlOverride" as="element(w:lvlOverride)?"
      select="key(
                'numbering-by-id', 
                @w:val,
                $root
              )/w:lvlOverride[@w:ilvl = $ilvl]"/>
    <xsl:apply-templates select="$lvl" mode="#current">
      <xsl:with-param name="numId" select="@w:val"/>
      <xsl:with-param name="start-override" select="for $so in (
                                                      $lvlOverride/w:startOverride/@w:val,
                                                      0[exists($lvlOverride)] (: this is subtle: if there is an empty override,
                                                                                 0 will be assumed as startOverride. Example:
                                                                                 Beuth 24739 :)
                                                    )[1]
                                                    return xs:integer($so)"/>
    </xsl:apply-templates>
  </xsl:template>

  <!-- It seems that numId = '0' is for lists without marker? -->
  <xsl:template match="w:numId[@w:val = '0']" mode="docx2hub:abstractNum" xml:id="manual-marker-aux-atts">
      <!-- Happens for <w:numId w:val="0"/>, when there is a manually set list counter -->
    <!-- Caution: key() seems to return empty sequence when this is invoked with saxon in mode {http://transpect.io/docx2hub}remove-redundant-run-atts -->
    <xsl:variable name="numPr-from-style" 
      select="key('docx2hub:style-by-role', ../../@role)/w:numPr" as="element(w:numPr)?"/>
    <xsl:variable name="abstractNum-from-style" 
      select="docx2hub:abstractNum-for-numPr($numPr-from-style)" as="element(w:abstractNum)?"/>
    <xsl:if test="$abstractNum-from-style">
      <!-- If the automatically generated numbering has been replaced with manually assigned numbering,
        we still must convey the abstract numbering that originally pertained to this paragraph.
        Otherwise the numbering will be reset at the next automatic numbering occasion for
        this abstractNum/ilvl. See @xml:id=('continue1', 'continue2') -->
      <xsl:attribute name="docx2hub:num-abstract" select="$abstractNum-from-style/@w:abstractNumId"/>
      <xsl:attribute name="docx2hub:num-signature" 
        select="string-join(($abstractNum-from-style/@w:abstractNumId, $numPr-from-style/w:ilvl[1]/@w:val), '_')"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:lvl" mode="docx2hub:abstractNum">
    <xsl:param name="numId" as="xs:string"/>
    <xsl:param name="start-override" as="xs:integer?"/>
    <xsl:variable name="restart" as="xs:integer?" select="(for $r in w:lvlRestart/@w:val 
                                                           return xs:integer($r),
                                                           0[$start-override])[last()]"/>
    <xsl:attribute name="docx2hub:num-signature" select="string-join((../@w:abstractNumId, @w:ilvl), '_')"/>
    <xsl:attribute name="docx2hub:num-abstract" select="../@w:abstractNumId"/>
    <xsl:attribute name="docx2hub:num-ilvl" select="@w:ilvl"/>
    <xsl:attribute name="docx2hub:num-id" select="$numId"/>
    <xsl:attribute name="docx2hub:num-lvlRestart" select="w:lvlRestart/@w:val"/>
    <xsl:if test="exists($restart)">
      <xsl:attribute name="docx2hub:num-restart-level" select="$restart"/>
    </xsl:if>
    <xsl:attribute name="docx2hub:num-restart-val" 
        select="($start-override, for $s in w:start/@w:val return xs:integer($s), 1)[1]"/>
    <xsl:if test="$start-override">
      <xsl:attribute name="docx2hub:num-start-override" select="$start-override"/>
    </xsl:if>
  </xsl:template>
  
  <!-- collateral (only the first in a row should trigger a reset) -->
  <xsl:template match="@docx2hub:num-signature[../preceding-sibling::*[1]
                                                                      [@docx2hub:num-signature = current()]
                                              ](:[  not sure about this 
                                                not(../@docx2hub:num-start-override)
                                              ]:)"
    mode="docx2hub:join-instrText-runs" xml:id="continue1">
    <!-- see @xml:id='manual-marker-aux-atts' -->
    <xsl:attribute name="docx2hub:num-continue" select="."/>
  </xsl:template>

  <xsl:template match="@docx2hub:num-signature[exists(../@docx2hub:num-restart-level)] (: should check whether 0 or a higher number :)
                                              [
                                                .. is (
                                                       key('docx2hub:num-signature', current())
                                                         [@docx2hub:num-id = current()/../@docx2hub:num-id]
                                                     )[1]
                                              ]" mode="docx2hub:join-instrText-runs" priority="2">
    <!-- the first of a numId that defines a start value override for this ilvl -->
    <xsl:copy/>
  </xsl:template>

  <xsl:template match="@docx2hub:num-signature" mode="docx2hub:join-instrText-runs">
    <xsl:variable name="context" as="element(w:p)" select=".."/>
    <xsl:variable name="last-same-signature" as="element(w:p)?" 
      select="(key('docx2hub:num-signature', current())[. &lt;&lt; $context])[last()]"/>
    <xsl:variable name="in-between" as="element(w:p)*"
      select="//w:p[if ($last-same-signature) 
                    then (. &gt;&gt; $last-same-signature)
                    else true()]
                   [. &lt;&lt; current()/..]"/>
    <xsl:variable name="same-abstract-in-between" as="element(w:p)*" 
      select="$in-between[@docx2hub:num-abstract = $context/@docx2hub:num-abstract]"/>
    <xsl:variable name="super-level-before" as="xs:boolean"
      select="some $p in $same-abstract-in-between satisfies 
              $p/@docx2hub:num-ilvl &lt; $context/@docx2hub:num-ilvl"/>
    <xsl:attribute name="docx2hub:num-super-level-before" select="$super-level-before"/>
    <xsl:attribute name="docx2hub:num-last-same-signature" select="exists($last-same-signature)"/>
    <xsl:choose>
      <xsl:when test="empty ($last-same-signature)
                      or 
                      $super-level-before">
        <xsl:copy/>
        <xsl:attribute name="docx2hub:num-restart-level" select="0"/>
      </xsl:when>
      <xsl:when test="(: $last-same-signature[not(@docx2hub:num-ilvl)] :) ../@docx2hub:num-start-override" xml:id="continue2">
        <!-- see #manual-marker-aux-atts -->
        <xsl:copy/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="docx2hub:num-continue" select="."/>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="empty($last-same-signature)">
      <!-- special case: sub levels of the same abstract def are before context; implicitly increase starting value by 1 
      So that it’s 1.1.1 instead of 0.0.1 if heading 3 is the first heading -->
      <xsl:if test="some $p in $same-abstract-in-between satisfies 
                    $p/@docx2hub:num-ilvl &gt; $context/@docx2hub:num-ilvl">
        <xsl:attribute name="docx2hub:num-initial-skip-increment" select="'1'"/>
      </xsl:if>
    </xsl:if>  
  </xsl:template>

  <xsl:function name="tr:insert-numbering" as="item()*">
    <xsl:param name="context" as="element(w:p)"/>
    <!-- Do we process lvlOverrides? -->
    
    <xsl:variable name="lvl" select="tr:get-lvl-of-numbering($context)" as="element(w:lvl)?"/>
    <xsl:choose>
      <xsl:when test="exists($lvl)">
        <xsl:if test="not($lvl/w:lvlText)">
          <xsl:call-template name="signal-error">
            <xsl:with-param name="error-code" select="'W2D_061'"/>
            <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
            <xsl:with-param name="hash">
              <value key="xpath">
                <xsl:value-of select="$lvl/@srcpath"/>
              </value>
              <value key="level">INT</value>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
        <xsl:variable name="style-atts" select="key('style-by-name', $context/@role, $context/root())/@*" as="attribute(*)*"/>
        <xsl:variable name="ad-hoc-atts" select="$context/@*" as="attribute(*)*"/>
        <xsl:variable name="pPr-from-numPr" as="attribute(*)*">
          <xsl:apply-templates mode="numbering" select="$lvl/w:pPr/@*">
            <xsl:with-param name="context" select="$context" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:variable>
        <xsl:variable name="rPr" as="attribute(*)*">
          <xsl:apply-templates mode="numbering" select="$lvl/w:rPr/@*">
            <xsl:with-param name="context" select="$context" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:variable>
        <xsl:variable name="immediate-first" as="attribute(*)*">
          <xsl:choose>
            <xsl:when test="$context/w:numPr">
              <xsl:sequence select="$style-atts[name() = $pPr-from-numPr/name()], $pPr-from-numPr"/>
            </xsl:when>
            <xsl:otherwise>
              <!-- the declaration priorities within $pPr (taking into account style inheritance) should have
                been sorted out when calculating $pPr during prop mapping -->
              <xsl:sequence select="$pPr-from-numPr, $style-atts[name() = $pPr-from-numPr/name()]"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:sequence select="$immediate-first, $ad-hoc-atts[name() = $pPr-from-numPr/name()]"/>
        <xsl:apply-templates select="$context/dbk:tabs" mode="wml-to-dbk"/>
        <phrase role="hub:identifier">
          <xsl:sequence select="$rPr, $style-atts[name() = $rPr/name()], $ad-hoc-atts[name() = $rPr/name()]"/>
          <xsl:if test="$rPr/self::attribute(docx2hub:map-from)">
            <!-- If the list marker character was in a mapped font, the replacement font should appear here. 
              Another unresolved issue might be: If the $style-atts @css:font-family is inherited from a
              w:basedOn style, and if the numPr are attached to the current style, then we will see the
              inherited font here instead of the numPr font. Not sure how Word behaves in that case. -->
            <xsl:sequence select="$rPr/self::attribute(docx2hub:map-from), $rPr/self::attribute(css:font-family)"/>
          </xsl:if>
          <xsl:value-of select="tr:get-identifier($context,$lvl)"/>
        </phrase>
        <tab/>
      </xsl:when>
      <xsl:otherwise>
        <!--KW 11.6.13: mode hart reingeschrieben wegen null pointer exception-->
        <xsl:apply-templates select="$context/dbk:tabs" mode="wml-to-dbk"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:function name="docx2hub:lvl-for-numPr" as="element(w:lvl)?">
    <xsl:param name="numPr" as="element(w:numPr)?"/>
    <xsl:sequence select="docx2hub:abstractNum-for-numPr($numPr)//w:lvl[@w:ilvl = $numPr/w:ilvl/@w:val]"/>
  </xsl:function>
  
  <xsl:function name="docx2hub:abstractNum-for-numPr" as="element(w:abstractNum)?">
    <xsl:param name="numPr" as="element(w:numPr)?"/>
    <xsl:sequence select="if (count(root($numPr)/*) = 1)
                          then
                            key(
                              'abstract-numbering-by-id', 
                              key(
                                'numbering-by-id', 
                                $numPr/w:numId/@w:val, 
                                root($numPr)
                              )/w:abstractNumId/@w:val,
                              root($numPr)
                            )
                          else ()"/>
  </xsl:function>

  <xsl:key name="abstractNum-by-styleLink" match="w:abstractNum" use="w:styleLink/@w:val"/>

  <xsl:function name="tr:get-lvl-of-numbering" as="element(w:lvl)?">
    <xsl:param name="context" as="element(w:p)"/>
    <!-- This function has not been migrated to make use of @docx2hub:num-… atts. Don’t know whether
      they may be exploited to make the function less verbose -->
    <!-- for-each: just to avoid an XTDE1270 which shouldn't happen when the 3-arg form of key() is invoked: -->
    <xsl:variable name="lvls" as="element(w:lvl)*">
      <xsl:for-each select="$context">
        <xsl:variable name="numPr" select="if ($context/w:numPr) 
                                           then $context/w:numPr 
                                           else ()"/>
        <xsl:variable name="numPr-from-pstyle" select="key('docx2hub:style-by-role', @role, root($context))[last()]/w:numPr" as="element(w:numPr)?"/>
        <xsl:variable name="lvl-for-numPr" as="element(w:lvl)?" select="docx2hub:lvl-for-numPr($numPr)"/>
        <xsl:variable name="abstractNum-for-numPr-from-pstyle" as="element(w:abstractNum)?" select="docx2hub:abstractNum-for-numPr($numPr-from-pstyle)"/>
        <xsl:variable name="lvl-for-numPr-from-pstyle" as="element(w:lvl)?" select="docx2hub:lvl-for-numPr($numPr-from-pstyle)"/>
        <xsl:sequence select="if ($numPr)
                              then if (exists($lvl-for-numPr)) 
                                then $lvl-for-numPr
                                else root($context)//w:abstractNum[
                                       w:styleLink/@w:val = docx2hub:abstractNum-for-numPr($numPr)/w:numStyleLink/@w:val
                                     ]/w:lvl[@w:ilvl = $numPr/w:ilvl/@w:val]
                              else if ($numPr-from-pstyle)
                                then if ($numPr-from-pstyle/w:ilvl/@w:val) 
                                  then if (exists($lvl-for-numPr-from-pstyle))
                                    then $lvl-for-numPr-from-pstyle
                                    else key(
                                           'abstractNum-by-styleLink', 
                                           $abstractNum-for-numPr-from-pstyle/w:numStyleLink/@w:val,
                                           root($context)
                                         )/w:lvl[@w:ilvl = $numPr-from-pstyle/w:ilvl/@w:val]
                                  else if ($context/@role and exists($abstractNum-for-numPr-from-pstyle/w:lvl[w:pStyle[@w:val = $context/@role]]))
                                    then $abstractNum-for-numPr-from-pstyle/w:lvl[w:pStyle[@w:val = $context/@role]]
                                    else if ($context/@role 
                                             and 
                                             exists(
                                               key(
                                                 'abstractNum-by-styleLink', 
                                                 $abstractNum-for-numPr-from-pstyle/w:numStyleLink/@w:val,
                                                 root($context)
                                               )/w:lvl[w:pStyle[@w:val = $context/@role]]
                                             )
                                            )
                                      then key(
                                             'abstractNum-by-styleLink', 
                                             $abstractNum-for-numPr-from-pstyle/w:numStyleLink/@w:val,
                                             root($context)
                                           )/w:lvl[w:pStyle[@w:val = $context/@role]]
                                      else $abstractNum-for-numPr-from-pstyle/w:lvl[@w:ilvl = '0']
                                else ()"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:sequence select="$lvls[last()]"/>
    <!--    Only last lvl chosen, because of errors. Check for multiple lvls has to be implemented   -->
  </xsl:function>
  
  <xsl:function name="tr:get-lvl-override" as="element(*)?">
    <xsl:param name="context" as="element(w:p)?"/>
    <xsl:choose>
      <xsl:when test="empty($context)">
        <xsl:sequence select="()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="numPr" select="$context/w:numPr" as="element(w:numPr)?"/>
        <xsl:variable name="style" as="element(w:numPr)?" 
          select="key('docx2hub:style-by-role', $context/@role, root($context))/w:numPr"/>
        <xsl:sequence select="if ($numPr)
                              then key('numbering-by-id', $numPr/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $numPr/w:ilvl/@w:val]
                              else if ($style)
                                   then key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $style/w:ilvl/@w:val]
                                   else ()"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:key name="docx2hub:num-signature" match="*[@docx2hub:num-signature]" use="@docx2hub:num-signature"/>
  <xsl:key name="docx2hub:num-signature-or-continue" match="*[@docx2hub:num-signature | @docx2hub:num-continue]" 
    use="@docx2hub:num-signature | @docx2hub:num-continue"/>

  <xsl:function name="tr:get-identifier" as="xs:string">
    <xsl:param name="context" as="element(w:p)"/>
    <xsl:param name="lvl" as="element(w:lvl)"/>
    
    <xsl:variable name="abstract-num-id" select="xs:double($lvl/ancestor::w:abstractNum/@w:abstractNumId)" as="xs:double"/>
    <xsl:variable name="lvl-to-use" select="if (exists(tr:get-lvl-override($context)/w:lvl)) 
                                            then tr:get-lvl-override($context)/w:lvl 
                                            else $lvl"/>
    <xsl:variable name="ilvl" select="xs:double($lvl-to-use/@w:ilvl)"/>
    
    <xsl:variable name="resolve-symbol-encoding">
      <element>
        <xsl:apply-templates select="$lvl-to-use/w:lvlText/@w:val" mode="wml-to-dbk"/>
      </element>
    </xsl:variable>
    <xsl:variable name="string" as="xs:string*">
      <xsl:choose>
        <xsl:when test="$resolve-symbol-encoding//@w:val">
          <xsl:analyze-string select="$lvl-to-use/w:lvlText/@w:val" regex="%(\d)">
            <xsl:matching-substring>
              <xsl:variable name="pattern-ilvl" as="xs:integer" select="xs:integer(regex-group(1)) - 1"/>
              <xsl:variable name="pattern-lvl" as="element(w:lvl)?" 
                select="$lvl/ancestor::w:abstractNum/w:lvl[@w:ilvl = $pattern-ilvl]"/>
              <xsl:variable name="context-for-counter" as="element(w:p)?">
                <xsl:choose>
                  <xsl:when test="$pattern-ilvl = $ilvl">
                    <xsl:sequence select="$context"/>
                  </xsl:when>
                  <xsl:when test="$pattern-ilvl gt $ilvl"/>
                  <xsl:otherwise>
                    <xsl:sequence select="(
                                            key(
                                              'docx2hub:num-signature-or-continue',
                                              string-join(($context/@docx2hub:num-abstract, string($pattern-ilvl)), '_'), 
                                              root($context)
                                            ) [not(w:numPr/w:numId/@w:val = '0')]
                                              [. &lt;&lt; $context]
                                          ) [last()]"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name="level-counter" as="xs:integer"
                select="if ($context-for-counter)
                        then tr:get-level-counter($context-for-counter, $pattern-lvl)
                        else $pattern-lvl/w:start/@w:val"/>
              <xsl:number value="$level-counter"
                          format="{tr:get-numbering-format($pattern-lvl/w:numFmt/@w:val, $lvl-to-use/w:lvlText/@w:val)}"/>
              <xsl:if
                test="$context/@srcpath = ('word/document.xml?xpath=/w:document[1]/w:body[1]/w:p[370]', 'word/document.xml?xpath=/w:document[1]/w:body[1]/w:p[399]')">
                <xsl:message
                  select="'AAAAAAAAAAAAAAA ', $level-counter, $context-for-counter"
                />
              </xsl:if>
            </xsl:matching-substring>
            <xsl:non-matching-substring>
              <xsl:value-of select="."/>
            </xsl:non-matching-substring>
          </xsl:analyze-string>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="$resolve-symbol-encoding//text()"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:sequence select="string-join($string,'')"/>
  </xsl:function>

  <xsl:function name="tr:get-level-counter" as="xs:integer">
    <xsl:param name="context" as="element(w:p)"/>
    <xsl:param name="lvl" as="element(w:lvl)"/>
    <xsl:variable name="start-of-relevant" as="element(w:p)?"
      select="if ($context/@docx2hub:num-signature)
              then $context
              else
                (
                  key(
                    'docx2hub:num-signature', 
                    ($context/@docx2hub:num-signature, $context/@docx2hub:num-continue), 
                    root($context)
                  )[. &lt;&lt; $context]
                )[last()]"/>
    <xsl:variable name="level-counter" as="xs:integer" 
      select="(for $s in $start-of-relevant/@docx2hub:num-restart-val return xs:integer($s), 1)[1] 
              + count(root($context)//w:p[. &gt;&gt; $start-of-relevant]
                                         [. &lt;&lt; $context]
                                         [@docx2hub:num-continue = $start-of-relevant/@docx2hub:num-signature]
                                         [not(w:numPr/w:numId/@w:val = '0')]
                     )
              + count($context[not(. is $start-of-relevant)])
              + count($start-of-relevant/@docx2hub:num-initial-skip-increment)"/>
    <xsl:sequence select="$level-counter"/>
  </xsl:function>

  <xsl:function name="tr:get-numbering-format" as="xs:string">
    <xsl:param name="format" as="xs:string"/>
    <xsl:param name="default" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="$format = 'lowerLetter'">a</xsl:when>
      <xsl:when test="$format = 'upperLetter'">A</xsl:when>
      <xsl:when test="$format = 'decimal'">1</xsl:when>
      <xsl:when test="$format = 'lowerRoman'">i</xsl:when>
      <xsl:when test="$format = 'upperRoman'">I</xsl:when>
      <xsl:when test="$format = 'bullet'">
        <xsl:value-of select="$default"/>
      </xsl:when>
      <xsl:when test="$format = 'none'">none</xsl:when><!--GR-->
      <xsl:otherwise>
        <!-- fallback: return 'none' (http://mantis.le-tex.de/mantis/view.php?id=13389#c36016) -->
        <xsl:text>none</xsl:text>
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_062'"/>
          <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
          <xsl:with-param name="hash">
            <value key="level">INT</value>
            <value key="info-text"><xsl:value-of select="$format"/>INT</value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:template match="@docx2hub:num-signature | @docx2hub:num-continue | @docx2hub:num-abstract | @docx2hub:num-id 
                       | @docx2hub:num-restart-val | @docx2hub:num-ilvl | @docx2hub:num-restart-level | @docx2hub:num-lvlRestart
                       | @docx2hub:num-super-level-before | @docx2hub:num-last-same-signature
                       | @docx2hub:num-initial-skip-increment | @docx2hub:num-start-override" mode="docx2hub:join-runs"/>
  
</xsl:stylesheet>