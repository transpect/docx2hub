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

  <!-- docx2hub:abstractNum is a collateral mode that is invoked from docx2hub:remove-redundant-run-atts (map-props.xsl) -->
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
    <xsl:variable name="lvlOverride" as="element(w:lvlOverride)?"
      select="key(
                'numbering-by-id', 
                @w:val,
                $root
              )/w:lvlOverride[@w:ilvl = $ilvl][last()]"/><!-- why can there be more than one? Are they guaranteed to have
                the same value? -->
    <xsl:variable name="lvl" as="element(w:lvl)?" 
      select="($abstractNum/w:lvl[@w:ilvl = $ilvl], $lvlOverride/w:lvl[@w:ilvl = $ilvl])[1]"/>
    <xsl:apply-templates select="$lvl" mode="#current">
      <xsl:with-param name="numId" select="@w:val"/>
      <xsl:with-param name="start-override" 
        select="for $so in (
                             $lvlOverride/w:startOverride/@w:val,
                             0[exists($lvlOverride[empty(*)])] (: this is subtle: if there is an empty override,
                                                                  0 will be assumed as startOverride. Example:
                                                                  Beuth 24739 :)
                           )[1]
                           return xs:integer($so)"/>
    </xsl:apply-templates>
  </xsl:template>

  <!-- docx2hub:abstractNum is a collateral mode that is invoked from docx2hub:remove-redundant-run-atts (map-props.xsl) -->
  <!-- numId = '0' is for lists without marker:
  17.9.19: “A value of 0 for the val attribute shall never be used to point to a numbering definition instance, and shall
instead only be used to designate the removal of numbering properties at a particular level in the style hierarchy
(typically via direct formatting).” -->
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
      <xsl:attribute name="docx2hub:num-disabled" select="'true'"/>
    </xsl:if>
  </xsl:template>

  <!-- docx2hub:abstractNum is a collateral mode that is invoked from docx2hub:remove-redundant-run-atts (map-props.xsl) -->
  <xsl:template match="w:lvl" mode="docx2hub:abstractNum">
    <xsl:param name="numId" as="xs:string"/>
    <xsl:param name="start-override" as="xs:integer?"/>
    <!-- “17.9.11 lvlRestart (Restart Numbering Level Symbol)
This element specifies a one-based index which determines when a numbering level should restart to its start
value (§17.9.26). A numbering level restarts when an instance of the specified numbering level, which shall be
higher (earlier than the this level) is used in the given document's contents.
If this element is omitted, the numbering level shall restart each time the previous numbering level is used. If
the specified level is higher than the current level, then this element shall be ignored. As well, a value of 0 shall
specify that this level shall never restart.” 

What follows in ISO/IEC 29500-1 (2006) are two examples, one with lvlRestart omitted and the other one with
lvlRestart = 0. We now try to construct an example where an ilvl=2 list should be reset at each ilvl=0 heading
but not at an ilvl=1 heading, as would be the default with omitted lvlRestart.
<w:lvl w:ilvl="3">
  <w:lvlRestart w:val="2">
</w:lvl>
This tells the ilvl=3 numbering to reset itself when an ilvl=1 (corresponds to lvlRestart=2) heading precedes
it, but not when an ilvl=2 heading precedes it.
    -->
    <xsl:variable name="restart-level" as="xs:integer?" select="for $r in w:lvlRestart/@w:val  
                                                                return xs:integer($r)"/>
    <xsl:attribute name="docx2hub:num-signature" select="string-join((../@w:abstractNumId, @w:ilvl), '_')"/>
    <xsl:attribute name="docx2hub:num-abstract" select="../@w:abstractNumId"/>
    <xsl:attribute name="docx2hub:num-ilvl" select="@w:ilvl"/>
    <xsl:attribute name="docx2hub:num-id" select="$numId"/>
    <xsl:variable name="restart-val" as="xs:integer"
      select="($start-override, for $s in w:start/@w:val return xs:integer($s), 0)[1]"/>
    <xsl:choose>
      <xsl:when test="$restart-level = 0">
        <xsl:attribute name="docx2hub:num-dont-restart" select="'yes'"/>
      </xsl:when>
      <xsl:when test="$restart-level &gt;= @w:ilvl">
        <!-- “If the specified level is higher than the current level, then this element shall be ignored.”
          What word does: It does not generate a number for this level. As I understood the spec, it says
          the w:lvlRestart element should be ignored, i.e., treated as if it wasn’t present. -->
        <xsl:attribute name="docx2hub:num-ignore-restart-level" select="'yes'"/>
      </xsl:when>
      <xsl:when test="exists($restart-level)">
        <!-- restart after preceding ilvl number (of the same abstractNumbering, I guess) lower than or equal as this value: --> 
        <xsl:attribute name="docx2hub:num-restart-after-ilvl" select="$restart-level - 1"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="docx2hub:num-restart-after-ilvl" select="@w:ilvl - 1"/>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:attribute name="docx2hub:num-restart-val" select="($restart-val, $start-override)[1]"/>
    <xsl:if test="$start-override">
      <xsl:attribute name="docx2hub:num-start-override" select="$start-override"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:p/@docx2hub:num-signature" mode="docx2hub:join-instrText-runs">
    <xsl:copy/>
    <xsl:variable name="counter-atts" as="attribute(*)*" 
      select="docx2hub:level-counter(.., ../@docx2hub:num-ilvl)"/>
    <xsl:if test="exists($counter-atts)">
      <xsl:sequence select="$counter-atts"/>
    </xsl:if>
  </xsl:template>

  <!-- recursive function that calculates the counter for the given ilvl based on the counter of the previous (non-disabled) 
    list paragraph. It outputs multiple attributes so that we can add more information to the intermediate debug file.
    Of the recursive invocation results, only the @docx2hub:num-counter-ilvl{$ilvl} attribute is used.
  -->
  <xsl:function name="docx2hub:level-counter" as="attribute(*)*">
    <xsl:param name="context" as="element(w:p)"/>
    <xsl:param name="ilvl" as="xs:integer?"/>
    <xsl:variable name="ilvl" as="xs:integer?" select="$context/@docx2hub:num-ilvl"/>
    <xsl:variable name="last-same-signature" as="element(w:p)?" 
      select="(key('docx2hub:num-signature', $context/@docx2hub:num-signature, root($context))
                 [. &lt;&lt; $context]
                 [not(@docx2hub:num-disabled = 'true')]
              )[last()]"/>
    <xsl:variable name="same-abstract-in-between" as="element(w:p)*"
      select="key('docx2hub:num-abstract', $context/@docx2hub:num-abstract, root($context))
                [if ($last-same-signature) 
                 then (. &gt;&gt; $last-same-signature)
                 else true()]
                [. &lt;&lt; $context]"/>
    <xsl:variable name="last-resetter" as="element(w:p)?" 
      select="($same-abstract-in-between[@docx2hub:num-ilvl &lt;= $context/@docx2hub:num-restart-after-ilvl],
               $context[@docx2hub:num-restart-val]
                       [@docx2hub:num-start-override]
                       [. is 
                          (//w:p[@docx2hub:num-id = $context/@docx2hub:num-id]
                                [not(@docx2hub:num-disabled = 'true')]
                           )[1] (: the first of a given numId may trigger a reset :)
                       ]
               )[last()]"/>
    <xsl:variable name="any-resetter" as="element(w:p)*" 
      select="key('docx2hub:num-abstract', $context/@docx2hub:num-abstract, root($context))
                [@docx2hub:num-ilvl &lt;= $context/@docx2hub:num-restart-after-ilvl or @docx2hub:num-ilvl = $context/@docx2hub:num-ilvl]
                [. &lt;&lt; $context]"/>
    <xsl:variable name="initial-sub-item" as="element(w:p)*"
      select="key('docx2hub:num-abstract', $context/@docx2hub:num-abstract, root($context))
                [empty($any-resetter) or (. &gt;&gt; $last-resetter)]
                [. &lt;&lt; $context]
                [@docx2hub:num-ilvl &gt; $context/@docx2hub:num-ilvl]
                [not(some $i 
                     in (key('docx2hub:num-abstract', $context/@docx2hub:num-abstract, root($context))[empty($any-resetter) or (. &gt;&gt; $last-resetter)][@docx2hub:num-ilvl &lt; $context/@docx2hub:num-ilvl]) 
                     satisfies (. &gt;&gt; $i))]"/>
    <xsl:if test="exists($ilvl)">
      <xsl:variable name="counter-name" as="xs:string" select="concat('docx2hub:num-counter-ilvl', $ilvl)"/>
      <xsl:variable name="counter" as="attribute(*)*">
        <xsl:choose>
          <xsl:when test="empty($last-same-signature) and exists($context/@docx2hub:num-dont-restart)">
            <xsl:attribute name="{$counter-name}" select="$context/@docx2hub:num-restart-val"/>
            <xsl:attribute name="docx2hub:num-debug-ilvl{$ilvl}-variant" select="'a'"/>
          </xsl:when>
          <xsl:when test="$context/@docx2hub:num-dont-restart and not($last-resetter is $context)">
            <xsl:attribute name="{$counter-name}" 
              select="for $a in docx2hub:level-counter($last-same-signature, $ilvl)[name() = $counter-name][. castable as xs:integer] 
                        return xs:integer($a)
                      + 1"/>
            <xsl:attribute name="docx2hub:num-debug-ilvl{$ilvl}-variant" select="'b'"/>
          </xsl:when>
          <xsl:when test="$last-resetter &gt;&gt; $last-same-signature">
            <xsl:attribute name="{$counter-name}" 
              select="(
                        for $rv in $context/@docx2hub:num-restart-val[. castable as xs:integer]
                        return xs:integer($rv),
                        4321
                      )[1]
                      + (1[exists($initial-sub-item)], 0)[1]"/>
            <xsl:attribute name="docx2hub:num-debug-ilvl{$ilvl}-variant" select="string-join(('c', '1'[exists($initial-sub-item)]), '')"/>
          </xsl:when>
          <xsl:when test="exists($last-same-signature)">
            <xsl:variable name="by-lookup" as="xs:integer?" 
              select="for $a in docx2hub:level-counter($last-same-signature, $ilvl)[name() = $counter-name][. castable as xs:integer] 
                        return xs:integer($a)"/>
            <xsl:choose>
              <xsl:when test="exists($by-lookup)">
                <xsl:attribute name="{$counter-name}" select="$by-lookup + 1"/>    
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name="docx2hub:num-debug-ilvl{$ilvl}-variant" select="'d'"/>
          </xsl:when>
          <xsl:when test="$context/@docx2hub:num-start-override">
            <!-- what about @docx2hub:num-restart-val? -->
            <xsl:attribute name="{$counter-name}" select="$context/@docx2hub:num-start-override"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="{$counter-name}" 
              select="(
                        for $rv in $context/@docx2hub:num-restart-val[. castable as xs:integer]
                        return xs:integer($rv),
                        9876
                      )[1]
                      + (1[exists($initial-sub-item)], 0)[1]"/>
            <xsl:attribute name="docx2hub:num-debug-ilvl{$ilvl}-variant" select="'e'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:sequence select="$counter"/>
      <xsl:if test="$debug = 'yes'">
        <xsl:attribute name="docx2hub:num-debug-last-same-signature-p" select="normalize-space($last-same-signature)"/>
        <xsl:attribute name="docx2hub:num-debug-last-resetter-p" select="normalize-space($last-resetter)"/>
        <xsl:attribute name="docx2hub:num-debug-initial-sub-item" select="exists($initial-sub-item)"/>
      </xsl:if>
    </xsl:if>
  </xsl:function>

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
        <!-- the declaration priorities within $pPr (taking into account style inheritance) should have
             been sorted out when calculating $pPr during prop mapping -->
        <xsl:variable name="immediate-first" as="attribute(*)*" 
                      select="if($context/w:numPr) then ($style-atts[name() = $pPr-from-numPr/name()], $pPr-from-numPr)
                              else ($pPr-from-numPr, $style-atts[name() = $pPr-from-numPr/name()])"/>
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
        <tab role="docx2hub:generated"/>
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
  
  <xsl:function name="docx2hub:lvl-for-numPr-and-ilvl" as="element(w:lvl)?">
    <xsl:param name="numPr" as="element(w:numPr)?"/>
    <xsl:param name="ilvl" as="xs:string?"/>
    <xsl:sequence select="docx2hub:abstractNum-for-numPr($numPr)//w:lvl[@w:ilvl = $ilvl]"/>
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
        <xsl:variable name="numPr" select="if ($context/w:numPr[node()]) 
                                           then $context/w:numPr[node()] 
                                           else ()"/>
        <xsl:variable name="numPr-from-pstyle" select="key('docx2hub:style-by-role', @role, root($context))[last()]/w:numPr" as="element(w:numPr)?"/>
        <!-- docx2hub:lvl-for-numPr-and-ilvl() is an attempt at making use of the new attributes in order
        to fix https://redmine.le-tex.de/issues/4224 -->
        <xsl:variable name="lvl-for-numPr" as="element(w:lvl)?" 
          select="(docx2hub:lvl-for-numPr-and-ilvl($numPr, $context/@docx2hub:num-ilvl), docx2hub:lvl-for-numPr($numPr))[1]"/>
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
                              then key('numbering-by-id', $numPr/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $numPr/w:ilvl/@w:val][last()]
                              else if ($style)
                                   then key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $style/w:ilvl/@w:val][last()]
                                   else ()"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:key name="docx2hub:num-abstract" match="*[@docx2hub:num-abstract]" use="@docx2hub:num-abstract"/>
  <xsl:key name="docx2hub:num-signature" match="*[@docx2hub:num-signature]" use="@docx2hub:num-signature"/>

  <xsl:function name="tr:get-identifier" as="xs:string">
    <xsl:param name="context" as="element(w:p)"/>
    <xsl:param name="lvl" as="element(w:lvl)"/>
    
    <xsl:variable name="abstract-num-id" select="xs:double($lvl/ancestor::w:abstractNum/@w:abstractNumId)" as="xs:double"/>
    <xsl:variable name="lvl-to-use" select="(tr:get-lvl-override($context)/w:lvl, $lvl)[1]" as="element(w:lvl)"/>
    <xsl:variable name="ilvl" select="xs:double($lvl-to-use/@w:ilvl)"/>
    
    <xsl:variable name="resolve-symbol-encoding" as="element()">
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
              
              <xsl:variable name="last-same-pattern-ilvl-signature" as="element(w:p)?" 
                select="(
                          key(
                            'docx2hub:num-signature', 
                            string-join(($context/@docx2hub:num-abstract, string($pattern-ilvl)), '_'), 
                            root($context)
                          )[. &lt;&lt; $context]
                        )[last()]"/>
              <xsl:variable name="same-abstract-in-between" as="element(w:p)*"
                select="key('docx2hub:num-abstract', $context/@docx2hub:num-abstract, root($context))
                          [if ($last-same-pattern-ilvl-signature) 
                           then (. &gt;&gt; $last-same-pattern-ilvl-signature)
                           else true()]
                          [. &lt;&lt; $context]"/>
              <xsl:variable name="last-resetter" as="element(w:p)?" 
                select="($same-abstract-in-between[@docx2hub:num-ilvl 
                                                   &lt;= 
                                                   $last-same-pattern-ilvl-signature/@docx2hub:num-restart-after-ilvl]
                        )[last()]"/>
              <xsl:variable name="context-for-counter" as="element(w:p)?">
                <xsl:choose>
                  <xsl:when test="$pattern-ilvl = $ilvl">
                    <xsl:sequence select="$context"/>
                  </xsl:when>
                  <xsl:when test="$pattern-ilvl gt $ilvl"/>
                  <xsl:otherwise>
                    <xsl:sequence select="(
                                            key(
                                              'docx2hub:num-signature',
                                              string-join(($context/@docx2hub:num-abstract, string($pattern-ilvl)), '_'), 
                                              root($context)
                                            ) [not(w:numPr/w:numId/@w:val = '0')]
                                              [. &lt;&lt; $context]
                                          ) [last()]"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name="level-counter" as="xs:integer?">
                <xsl:choose>
                  <xsl:when test="exists($last-resetter)
                                  and 
                                  not($context-for-counter is $context)">
                    <!-- 2.1.4 preceded by 2. The pattern-ilvl=1 resetter is closer 
                      than any same-ilvl para. Therefore use the start value. -->
                    <xsl:sequence select="$pattern-lvl/w:start/@w:val"></xsl:sequence>
                  </xsl:when>
                  <xsl:when test="exists($context-for-counter)
                                  and
                                  $context-for-counter/@*[name() = concat('docx2hub:num-counter-ilvl', $pattern-ilvl)]
                                  castable as xs:integer">
                    <xsl:sequence select="xs:integer($context-for-counter/@*[name() = concat('docx2hub:num-counter-ilvl', $pattern-ilvl)])"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:sequence select="$pattern-lvl/w:start/@w:val"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:if test="empty($level-counter)">
                <xsl:message select="'Empty level counter for', $pattern-ilvl, ':', $context"/>
              </xsl:if>
              <xsl:variable name="provisional-number">
                <xsl:number value="($level-counter, 9999)[1]"
                            format="{tr:get-numbering-format($pattern-lvl/w:numFmt/@w:val, $lvl-to-use/w:lvlText/@w:val)}"/>
              </xsl:variable>
              <xsl:variable name="cardinality" select="if (matches($provisional-number,'^\*†‡§[0-9]+\*†‡§$')) then xs:integer(replace($provisional-number, '^\*†‡§([0-9]+)\*†‡§$', '$1')) else 0"/>
              <xsl:value-of select="if (matches($provisional-number,'^\*†‡§[0-9]+\*†‡§$')) 
                                    then string-join((for $i 
                                                      in (1 to xs:integer(ceiling($cardinality div 4))) 
                                                      return substring($provisional-number,if (($cardinality mod 4) ne 0) then ($cardinality mod 4) else 4,1)),'') 
                                    else if (matches($provisional-number,'^a[a-z]$')) 
                                         then replace($provisional-number,'^a([a-z])$','$1$1')
                                         else $provisional-number"/>
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

  <xsl:function name="tr:get-numbering-format" as="xs:string">
    <xsl:param name="format" as="xs:string"/>
    <xsl:param name="default" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="$format = ('lowerLetter','lower-letter')">a</xsl:when>
      <xsl:when test="$format = ('upperLetter','upper-letter')">A</xsl:when>
      <xsl:when test="$format = ('decimal', 'ordinal')">1</xsl:when>
      <xsl:when test="$format = 'decimalZero'">01</xsl:when>
      <xsl:when test="$format = ('lowerRoman','lower-roman')">i</xsl:when>
      <xsl:when test="$format = ('upperRoman','upper-roman')">I</xsl:when>
      <xsl:when test="$format = 'bullet'">
        <xsl:value-of select="$default"/>
      </xsl:when>
      <xsl:when test="$format = 'chicago'">*†‡§</xsl:when>
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
  
  <xsl:template match="@*[starts-with(name(), 'docx2hub:num-')]" mode="docx2hub:join-runs"/>
  
</xsl:stylesheet>
