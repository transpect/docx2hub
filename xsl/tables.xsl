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
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0"
  exclude-result-prefixes = "w o v wx dbk pkg r rel word200x exsl saxon fn tr">

  <xsl:template match="w:tbl" mode="wml-to-dbk">
    <xsl:variable name="styledef" as="element(css:rule)?" select="key('docx2hub:style-by-role', w:tblPr/@role)"/>
    <informaltable>
      <xsl:variable name="collapse" select="(w:tblPr/@docx2hub:generated-tblCellSpacing/'separate','collapse')[1]" as="xs:string"/>
      <xsl:variable name="outer-border-set" select="
        empty(w:tr[1]/w:tc/w:tcPr/w:top[matches(@w:val, 'nil|none')][1]),
        empty(w:tr/w:tc[last()]/w:tcPr/w:right[matches(@w:val, 'nil|none')][1]),
        empty(w:tr[last()]/w:tc/w:tcPr/w:bottom[matches(@w:val, 'nil|none')][1]),
        empty(w:tr/w:tc[1]/w:tcPr/w:left[matches(@w:val, 'nil|none')][1])
      " as="xs:boolean+"/>
      <xsl:attribute name="css:border-collapse" select="$collapse"/>
      <xsl:apply-templates select="w:tblPr/@role,
                                   (
                                    $styledef/w:tblPr/@css:*[matches(name(.), '(border-(top|right|bottom|left)-(style|color|width)|background-color|margin-(left|right)|text-align)')], 
                                    w:tblPr/@css:*[matches(name(.), '(border-(top|right|bottom|left)-(style|color|width)|background-color|margin-(left|right)|text-align)')]
                                   )[$collapse = 'separate' or (matches(., 'background-color|margin-(left|right)|text-align')) or
                                    (
                                      (matches(name(), 'top') and $outer-border-set[1])
                                      or (matches(name(), 'right') and $outer-border-set[2])
                                      or (matches(name(), 'bottom') and $outer-border-set[3])
                                      or (matches(name(), 'left') and $outer-border-set[4])
                                    )],
                                   $styledef/w:tblPr/w:tblW, 
                                   w:tblPr/w:tblW,
                                   w:tblPr/@css:table-layout,
                                   @srcpath, 
                                   @css:orientation" mode="#current"/>
      <tgroup>
        <xsl:attribute name="cols" select="(
                                             count(w:tblGrid/w:gridCol),
                                             max(for $row in w:tr return count($row/w:tc)) (: fallback if w:tblGrid is missing :) 
                                           )[not(. = 0)][1]"/>
        <xsl:variable name="table-adhoc-borders" select="w:tblPr/@css:*[matches(name(.), '(border-(top|right|bottom|left)-(style|color|width))')]"/>
        <xsl:variable name="insideH-width" as="xs:string?" 
          select="($styledef/w:tblPr/@css:border-insideH-width, w:tblPr/@css:border-insideH-width)[last()]"/>
        <xsl:variable name="insideV-width" as="xs:string?" 
          select="($styledef/w:tblPr/@css:border-insideV-width, w:tblPr/@css:border-insideV-width)[last()]"/>
        <xsl:if test="exists($insideH-width) and not($insideH-width = '0pt')">
          <xsl:attribute name="rowsep" select="'1'"/>
        </xsl:if>
        <xsl:if test="exists($insideV-width) and not($insideV-width = '0pt')">
          <xsl:attribute name="colsep" select="'1'"/>
        </xsl:if>
        <xsl:apply-templates select="w:tblGrid" mode="colspec"/>
        <xsl:variable name="every-row-is-a-header" as="xs:boolean"
          select="every $tr in w:tr satisfies $tr[docx2hub:is-tableheader-row(.)]"/>
        <xsl:if test="w:tr[docx2hub:is-tableheader-row(.)] and 
                      not($every-row-is-a-header)">
          <thead>
            <xsl:apply-templates select="w:tr[docx2hub:is-tableheader-row(.) 
                                              and (position() eq 1 or preceding-sibling::*[1][self::w:tr]/docx2hub:is-tableheader-row(.))]" mode="tables">
              <xsl:with-param name="cols" select="count(w:tblGrid/w:gridCol)" tunnel="yes"/>
              <xsl:with-param name="width" select="w:tblPr/w:tblW/@w:w" tunnel="yes"/>
              <xsl:with-param name="col-widths" select="(for $x in w:tblGrid/w:gridCol return $x/@w:w)" tunnel="yes"/>
              <xsl:with-param name="table-borders" select="$table-adhoc-borders" as="attribute(*)*" tunnel="yes"/>
            </xsl:apply-templates>
          </thead>
        </xsl:if>
        <tbody>
          <xsl:apply-templates mode="tables"
            select="if($every-row-is-a-header) then * except w:tblPr
                    else * except (w:tblPr union w:tr[docx2hub:is-tableheader-row(.)
                                                      and (position() eq 1 or preceding-sibling::*[1][self::w:tr]/docx2hub:is-tableheader-row(.))])">
            <xsl:with-param name="cols" select="count(w:tblGrid/w:gridCol)" tunnel="yes"/>
            <xsl:with-param name="width" select="w:tblPr/w:tblW/@w:w" tunnel="yes"/>
            <xsl:with-param name="col-widths" select="(for $x in w:tblGrid/w:gridCol return $x/@w:w)" tunnel="yes"/>
            <xsl:with-param name="table-borders" select="$table-adhoc-borders" as="attribute(*)*" tunnel="yes"/>
          </xsl:apply-templates>
        </tbody>
      </tgroup>
    </informaltable>
  </xsl:template>

  <xsl:function name="docx2hub:is-tableheader-row" as="xs:boolean">
    <xsl:param name="row" as="element(w:tr)"/>
    <xsl:sequence select="$row/w:trPr/w:tblHeader or 
                          $row/w:tblHeader or
                          $row[w:tc/w:tcPr/w:vMerge[not(@w:val) or @w:val eq 'continue']]
                              [preceding-sibling::w:tr[not(docx2hub:is-blind-vmerged-row(.))][1]/self::w:tr[docx2hub:is-tableheader-row(.)]]"/>
  </xsl:function>

  <xsl:template match="w:tblGridChange" mode="colspec"/>

  <xsl:template match="w:tblPr/@css:background-color" priority="10" mode="wml-to-dbk">
    <xsl:sequence select="."/>
  </xsl:template>
  
  <xsl:template match="w:tblGrid" mode="colspec">
    <xsl:apply-templates mode="colspec"/>
  </xsl:template>
  
  <xsl:template match="w:tblPr/w:tblW[@w:type eq 'auto']" mode="wml-to-dbk" priority="2"/>

  <xsl:template match="w:tblPr/w:tblW[@w:type eq 'auto'][../w:tblLayout[@w:type eq 'fixed']]" mode="wml-to-dbk" priority="2.5">
      <!-- this will give a hint to the hub2docx converter that it is a auto-but-fixed-width table,
      i.e. that the cells’ widths will be preserved 
    --> 
    <xsl:attribute name="css:width" select="'auto'"/>
  </xsl:template>
  
  <xsl:template match="w:tblPr/w:tblW[@w:type eq 'dxa']" mode="wml-to-dbk" priority="2">
    <xsl:attribute name="css:width" select="docx2hub:pt-length(@w:w)"/>
  </xsl:template>
  
  <xsl:template match="w:tblPr/w:tblW[@w:type eq 'pct']" mode="wml-to-dbk" priority="2">
    <!-- Percentage with a percent sign or integer value for which 5000 = 100%.
    The latter is not in the spec but here: http://officeopenxml.com/WPtableWidth.php-->
    <xsl:attribute name="css:width" select="if(ends-with(@w:w, '%'))
                                            then @w:w
                                            else if (@w:w castable as xs:integer)
                                                 then concat(round(xs:double(@w:w) * 0.2) * 0.1, '%')
                                                 else @w:w"/>
  </xsl:template>
  
  <xsl:template match="w:tblPr/w:tblW" mode="wml-to-dbk">
    <!-- validation will complain about a unitless value --> 
    <xsl:attribute name="css:width" select="@w:w"/>
  </xsl:template>
  
  <xsl:template match="w:tblGrid/w:gridCol" mode="colspec">
    <xsl:variable name="pos" select="index-of(for $elt in ../* return generate-id($elt), generate-id())" as="xs:integer"/>
    <xsl:variable name="relative-table-width" as="xs:string?" 
      select="if (every $type in parent::*/following-sibling::w:tr/descendant::w:tcW/@w:type satisfies ($type eq 'pct')) 
              then 'percent' else ''"/>
    <xsl:variable name="colwidth-sum" as="xs:double" select="sum(parent::*/w:gridCol/@w:w)"/>
    <xsl:variable name="current-width" as="xs:integer" select="xs:integer(@w:w)"/>
    <colspec>
      <xsl:attribute name="colnum">
        <xsl:value-of select="$pos"/>
      </xsl:attribute>
      <xsl:attribute name="colname">
        <xsl:value-of select="concat('col', $pos)"/>
      </xsl:attribute>
      <xsl:if test="@w:w != ''">
        <xsl:choose>
          <xsl:when test="$relative-table-width eq 'percent'">
            <xsl:attribute name="colwidth" 
              select="concat(format-number(round(10000 * $current-width div $colwidth-sum) * 0.01, '#.00'), '%')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="colwidth" select="docx2hub:twips2mm(@w:w)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </colspec>
  </xsl:template>

  <xsl:template match="w:tblGrid" mode="tables"/>

  <xsl:template match="w:tblPr" mode="tables">
    <xsl:apply-templates select="@css:*[not(contains(local-name(), 'inside'))]" mode="#current"/>
  </xsl:template>

  <xsl:template match="w:commentRangeEnd | w:commentRangeStart | w:bookmarkStart | w:bookmarkEnd" mode="tables"/>

  <xsl:template match="*" mode="tables">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_020'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">tables</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="processing-instruction()" mode="tables">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_023'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">tables</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="comment()" mode="tables">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_022'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">tables</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:tr[docx2hub:is-blind-vmerged-row(.)]" mode="tables"/>
  
  <xsl:function name="docx2hub:is-blind-vmerged-row" as="xs:boolean">
    <xsl:param name="row" as="element(w:tr)"/>
    <xsl:sequence select="every $tc in $row/w:tc satisfies $tc[docx2hub:is-blind-vmerged-cell(.)]"/>
  </xsl:function>

  <xsl:template match="w:tr[not(docx2hub:is-blind-vmerged-row(.))]" mode="tables">
    <row>
      <xsl:apply-templates select="@css:*, w:tblPrEx/@css:background-color, @xml:lang, @srcpath" mode="wml-to-dbk"/>
      <xsl:apply-templates select="@w:fill-cells-before" mode="wml-to-dbk"/>
      <xsl:apply-templates select="w:tc" mode="#current">
        <!-- Despite its name, row-overrides will also contain the un-overridden insideV and insideH settings of tblPr, if available.-->
        <xsl:with-param name="row-overrides" as="attribute(*)*" 
          select="ancestor::w:tbl[1]/w:tblPr/@css:*[matches(local-name(), '(inside[HV]|padding)')], w:tblPrEx/(@* except @css:background-color)"/>
        <xsl:with-param name="is-first-row-in-group" select="tr:node-index-of(../w:tr, .) = 1"  tunnel="yes"/>
        <xsl:with-param name="is-last-row-in-group" select="tr:node-index-of(../w:tr, .) = count(../w:tr)"  tunnel="yes"/>
      </xsl:apply-templates>
      <xsl:apply-templates select="@w:fill-cells-after" mode="wml-to-dbk"/>
    </row>
  </xsl:template>
  
  <xsl:template match="@w:fill-width-before | @w:fill-width-after" mode="wml-to-dbk" priority="10"/>
  
  <xsl:template match="@w:fill-cells-before" mode="wml-to-dbk" priority="10">
    <!--<xsl:param name="col-widths" as="xs:integer+" tunnel="yes"/>
    <xsl:for-each select="1 to xs:integer(.)">
      <entry role="hub:fill-grid" css:width="{docx2hub:twips2mm($col-widths[current()])}">
        <para/>
      </entry>
    </xsl:for-each>  -->
  </xsl:template>
  
  <xsl:template match="@w:fill-cells-after" mode="wml-to-dbk" priority="10">
    <!--<xsl:param name="col-widths" as="xs:integer+" tunnel="yes"/>
    <xsl:param name="cols" as="xs:integer" tunnel="yes"/>
    <xsl:for-each select="($cols - xs:integer(.) + 1) to $cols">
      <entry role="hub:fill-grid" css:width="{docx2hub:twips2mm($col-widths[current()])}">
        <para/>
      </entry>
    </xsl:for-each>-->
  </xsl:template>

  <xsl:template match="w:trPr" mode="tables"/>

  <xsl:template match="w:tcPr" mode="tables"/>

  <xsl:template match="w:tblPrEx" mode="tables"/>

  <xsl:template match="w:p" mode="tables">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>

  <xsl:template match="w:tc[not(docx2hub:is-blind-vmerged-cell(.))]" mode="tables">
    <xsl:param name="is-first-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="col-widths" tunnel="yes" as="xs:integer*"/>
    <xsl:param name="row-overrides" as="attribute(*)*"/>
    <xsl:param name="table-borders" as="attribute(*)*" tunnel="yes"/>
    <xsl:element name="entry">
<!--      <xsl:copy-of select="ancestor::w:tbl[1]/w:tblPr/@css:*[not(matches(local-name(), '^(border|background-color|width)'))]"/>-->
      <xsl:sequence select="ancestor::w:tr[1]/@css:*[not(matches(local-name(), '^(background-color|(min-)?height|width)'))]"/>
      <xsl:if test="w:p[@css:text-align] and (every $t in w:p satisfies ($t[@css:text-align] and $t/@css:text-align=w:p[1]/@css:text-align))">
        <xsl:sequence select="w:p[1]/@css:text-align"/>  
      </xsl:if>
      <xsl:variable name="is-first-cell" select="tr:node-index-of(../w:tc, .) = 1" as="xs:boolean"/>
      <xsl:variable name="is-last-cell" select="tr:node-index-of(../w:tc, .) = count(../w:tc)" as="xs:boolean"/>
      <xsl:variable name="is-first-row-in-group" select="tr:node-index-of(../../w:tr, ..) = 1" as="xs:boolean"/>
      <xsl:variable name="is-last-row-in-group" select="tr:node-index-of(../../w:tr, ..) = count(../../w:tr)" as="xs:boolean"/>
      <xsl:apply-templates select="$row-overrides, @*" mode="wml-to-dbk">
        <xsl:with-param name="is-first-cell" select="$is-first-cell" tunnel="yes"/>
        <xsl:with-param name="is-last-cell" select="$is-last-cell" tunnel="yes"/>
        <xsl:with-param name="is-first-row-in-group" select="$is-first-row-in-group" tunnel="yes"/>
        <xsl:with-param name="is-last-row-in-group" select="$is-last-row-in-group" tunnel="yes"/>
      </xsl:apply-templates>
      <xsl:apply-templates select="$table-borders" mode="copy-borders">
        <xsl:with-param name="is-first-cell" select="$is-first-cell" tunnel="yes"/>
        <xsl:with-param name="is-last-cell" select="$is-last-cell" tunnel="yes"/>
        <xsl:with-param name="is-first-row-in-group" select="$is-first-row-in-group" tunnel="yes"/>
        <xsl:with-param name="is-last-row-in-group" select="$is-last-row-in-group" tunnel="yes"/>
      </xsl:apply-templates>
      <xsl:apply-templates select="@*" mode="wml-to-dbk">
        <xsl:with-param name="is-first-cell" select="$is-first-cell" tunnel="yes"/>
        <xsl:with-param name="is-last-cell" select="$is-last-cell" tunnel="yes"/>
        <xsl:with-param name="is-first-row-in-group" select="$is-first-row-in-group" tunnel="yes"/>
        <xsl:with-param name="is-last-row-in-group" select="$is-last-row-in-group" tunnel="yes"/>
      </xsl:apply-templates>
      <!-- Process any spans -->
      <xsl:call-template name="cell.span"/>
      <xsl:variable name="morerows" as="attribute()*">
        <xsl:call-template name="cell.morerows"/>
      </xsl:variable>
      <xsl:variable name="morerows-value" as="xs:integer">
        <xsl:value-of select="if (empty($morerows)) then 0 else $morerows"/>
      </xsl:variable>
      <xsl:if test="$morerows-value gt 0">
        <xsl:variable name="pos" select="docx2hub:get-tc-grid-position(ancestor::w:tr[1], .)"/>
        <xsl:variable name="bottom-border-row"
          select="ancestor::w:tr[1]/following-sibling::w:tr[position() = $morerows-value]" as="element(w:tr)"/>
        <xsl:variable name="bottom-border-cell"
          select="docx2hub:get-tc-by-grid-position($bottom-border-row, $pos)"
          as="element(w:tc)"/>
        <xsl:apply-templates select="$bottom-border-cell/@css:*[matches(name(), 'border\-bottom\-')]" mode="wml-to-dbk">
          <xsl:with-param name="is-first-cell" select="$is-first-cell" tunnel="yes"/>
          <xsl:with-param name="is-last-cell" select="$is-last-cell" tunnel="yes"/>
          <xsl:with-param name="is-first-row-in-group" select="$is-first-row-in-group" tunnel="yes"/>
          <xsl:with-param name="is-last-row-in-group" select="$is-last-row-in-group" tunnel="yes"/>
        </xsl:apply-templates>
      </xsl:if>
      <xsl:sequence select="$morerows"/>
      <!-- will be treated by map-props.xsl:
      <xsl:call-template name="cell.style"/>
      <xsl:call-template name="cell.align"/>
      -->
      <xsl:if test="not(preceding-sibling::w:tc) and 
                    parent::w:tr/preceding-sibling::*[1][self::w:commentRangeStart or self::w:bookmarkStart or self::w:bookmarkEnd]">
        <xsl:apply-templates select="parent::w:tr/preceding-sibling::*[1]
                                                                      [self::w:commentRangeStart or 
                                                                       self::w:bookmarkStart or 
                                                                       self::w:bookmarkEnd]" 
                             mode="wml-to-dbk"/>
      </xsl:if>
      <xsl:apply-templates select="*" mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <xsl:function name="docx2hub:get-tc-grid-position" as="xs:integer">
    <xsl:param name="row" as="element(w:tr)"/>
    <xsl:param name="cell" as="element(w:tc)"/>
    <xsl:variable name="id" select="generate-id($cell)"/>
    <xsl:variable name="preceding-tc" select="$row/w:tc[generate-id() = $id]/preceding-sibling::w:tc" as="node()*"/>
    <xsl:sequence select="
      1 +
      xs:integer(sum($preceding-tc/w:tcPr/w:gridSpan/@w:val))
      +
      count($preceding-tc[not(w:tcPr/w:gridSpan)])
    "/>
  </xsl:function>
  
  <xsl:function name="docx2hub:get-tc-by-grid-position" as="element(w:tc)">
    <xsl:param name="row" as="element(w:tr)"/>
    <xsl:param name="position" as="xs:integer"/>
    <!--
      example:
      | a | a | a | b | c1 |
      | d | d | e | b | c2 |
      | f | g | e | h | i  |
      for position 5 (c1), we need to return c2
      => account for preceding grid spans (like d)
    -->
    <xsl:sequence select="$row/w:tc[docx2hub:get-tc-grid-position($row, .) = $position]"/>
  </xsl:function>
<!--
  <xsl:template match="@css:border-insideH-style | @css:border-insideH-width | @css:border-insideH-color" mode="wml-to-dbk" priority="10">
    <xsl:param name="is-first-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:if test="not($is-first-row-in-group)">
      <xsl:attribute name="{replace(name(), 'insideH', 'top')}" select="."/>
    </xsl:if>
    <xsl:if test="not($is-last-row-in-group)">
      <xsl:attribute name="{replace(name(), 'insideH', 'bottom')}" select="."/>
    </xsl:if>
  </xsl:template>-->
  

  <xsl:template match="@css:border-insideH-style | @css:border-insideH-width | @css:border-insideH-color" mode="wml-to-dbk" priority="10">
      <xsl:attribute name="{replace(name(), 'insideH', 'top')}" select="."/>
      <xsl:attribute name="{replace(name(), 'insideH', 'bottom')}" select="."/>
  </xsl:template>

  <xsl:template match="@*[matches(local-name(), 'border-top')]" mode="copy-borders" priority="10">
    <xsl:param name="is-first-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:if test="$is-first-row-in-group">
      <xsl:copy/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="@*[matches(local-name(), 'border-bottom')]" mode="copy-borders" priority="10">
    <xsl:param name="is-first-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:if test="$is-last-row-in-group">
      <xsl:copy/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="@*[matches(local-name(), 'border-left')]" mode="copy-borders" priority="10">
    <xsl:param name="is-first-cell" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-cell" as="xs:boolean?" tunnel="yes"/>
     <xsl:if test="$is-first-cell">
      <xsl:copy/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="@*[matches(local-name(), 'border-right')]" mode="copy-borders" priority="10">
    <xsl:param name="is-first-cell" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-cell" as="xs:boolean?" tunnel="yes"/>
     <xsl:if test="$is-last-cell">
      <xsl:copy/>
    </xsl:if>
  </xsl:template>

  <!--<xsl:template match="@css:border-insideV-style | @css:border-insideV-width | @css:border-insideV-color" mode="wml-to-dbk" priority="10">
    <xsl:param name="is-first-cell" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-cell" as="xs:boolean?" tunnel="yes"/>
    <xsl:if test="not($is-first-cell)">
      <xsl:attribute name="{replace(name(), 'insideV', 'left')}" select="."/>
    </xsl:if>
    <xsl:if test="not($is-last-cell)">
      <xsl:attribute name="{replace(name(), 'insideV', 'right')}" select="."/>
    </xsl:if>
  </xsl:template>-->
  
  <xsl:template match="@css:border-insideV-style | @css:border-insideV-width | @css:border-insideV-color" mode="wml-to-dbk" priority="10">
      <xsl:attribute name="{replace(name(), 'insideV', 'left')}" select="."/>
      <xsl:attribute name="{replace(name(), 'insideV', 'right')}" select="."/>
  </xsl:template>

  <xsl:template name="cell.align">
    <!-- horizontale Ausrichtung des ersten Absatzes -->
    <xsl:variable name="align" select="if (exists(w:p/w:pPr/w:jc/@w:val)) then (w:p/w:pPr/w:jc/@w:val)[1] else 'left'"/>
    <!-- vertikale Ausrichtung -->
    <xsl:variable name="v-align" select="w:tcPr/w:vAlign/@w:val"/>
    <xsl:choose>
      <xsl:when test="$align = 'left'">
        <xsl:attribute name="css:text-align" select="'left'"/>
      </xsl:when>
      <xsl:when test="$align = 'right'">
        <xsl:attribute name="css:text-align" select="'right'"/>
      </xsl:when>
      <xsl:when test="$align = 'both'">
        <xsl:attribute name="css:text-align" select="'justify'"/>
      </xsl:when>
      <xsl:when test="$align = 'center'">
        <xsl:attribute name="css:text-align" select="'center'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="css:text-align" select="'left'"/>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="$v-align = ('both', 'center')">
        <xsl:attribute name="css:vertical-align" select="'middle'"/>        
      </xsl:when>
      <xsl:when test="$v-align = 'top'">
        <xsl:attribute name="css:vertical-align" select="'top'"/>        
      </xsl:when>
      <xsl:when test="$v-align = 'bottom'">
        <xsl:attribute name="css:vertical-align" select="'bottom'"/>        
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="css:vertical-align" select="'top'"/>        
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="tr:check-border-hierarchy" as="node()?">
    <xsl:param name="borders" as="node()?"/>
    <xsl:param name="name" as="xs:string"/>
    <xsl:param name="flag" as="xs:boolean"/>
    <xsl:param name="alternative" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$flag">
        <xsl:sequence select="$borders/*[name() = $name]"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$borders/*[name() = $alternative]">
            <xsl:element name="{$name}">
              <xsl:sequence select="$borders/*[name() = $alternative]/@*"/>
            </xsl:element>
          </xsl:when>
          <xsl:otherwise>
            <xsl:sequence select="$borders/*[name() = $name]"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>


  <xsl:function name="tr:write-table-style" as="xs:string">
    <xsl:param name="border" as="node()"/>
    <xsl:sequence select="string-join(for $x in $border/node() return
                          (for $y in $x/@*[name() = ('w:sz', 'w:val', 'w:color')]
                          return concat('border-', local-name($x), tr:translate-border-attribs($y))), '; ')"/>
  </xsl:function>

  <xsl:function name="tr:translate-border-attribs" as="xs:string">
    <xsl:param name="attrib" as="attribute()"/>
    <xsl:choose>
      <xsl:when test="name($attrib) = 'w:sz'">
        <xsl:value-of select="concat('-width: ', $attrib div 8, 'pt')"/>
      </xsl:when>
      <xsl:when test="name($attrib) = 'w:val'">
        <xsl:value-of select="concat('-style: ', tr:translate-border-style($attrib))"/>
      </xsl:when>
      <xsl:when test="name($attrib) = 'w:color'">
        <xsl:value-of select="concat('-color: ', $attrib)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="''"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="tr:translate-border-style" as="xs:string">
    <xsl:param name="text" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$text = 'nil'">none</xsl:when>
      <xsl:when test="$text = 'single'">solid</xsl:when>
      <xsl:otherwise>
        <!--
        <xsl:message>ACHTUNG: border-style kann nicht bestimmt werden: <xsl:value-of select="$text"/></xsl:message>
        -->
        <xsl:value-of select="'solid'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>


  <xsl:template match="w:tc[docx2hub:is-blind-vmerged-cell(.)]" mode="tables"/>

  <xsl:template name="cell.span">
    <xsl:variable name="span" select="0 + (if (w:tcPr/w:gridSpan/@w:val) then w:tcPr/w:gridSpan/@w:val - 1 else 0)
                                      + (if (not(following-sibling::w:tc) and ../@w:fill-cells-after) then ../@w:fill-cells-after else 0) 
                                      + (if (not(preceding-sibling::w:tc) and ../@w:fill-cells-before) then ../@w:fill-cells-before else 0)"/>
    <xsl:variable name="colstart" select="tr:colcount(1, .) + (if (preceding-sibling::w:tc and ../@w:fill-cells-before) then ../@w:fill-cells-before else 0)" />
    <xsl:choose>
      <xsl:when test="$span &gt; 0">
        <!-- Get the current cell number -->
        <xsl:attribute name="namest">
          <xsl:value-of  select="concat('col', $colstart)"/>
        </xsl:attribute>
        <xsl:attribute name="nameend">
          <xsl:value-of  select="concat('col', $colstart + $span)"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="colname">
          <xsl:value-of  select="concat('col', $colstart)"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="docx2hub:is-blind-vmerged-cell" as="xs:boolean">
    <xsl:param name="cell" as="element(w:tc)"/>
    <xsl:sequence select="exists($cell/w:tcPr/w:vMerge)
                          and
                          (every $att in $cell/w:tcPr/w:vMerge/@* satisfies (
                             $att/self::attribute(srcpath) or 
                             $att/self::attribute(w:val)[. = 'continue'] or 
                             $att/namespace-uri() eq 'http://www.w3.org/1996/css'
                          ))"/>
  </xsl:function>

  <xsl:template name="cell.morerows">
    <xsl:if test="w:tcPr/w:vMerge/@w:val = 'restart'">
      <xsl:variable name="is-thead-tr" as="xs:boolean" 
        select="exists(parent::w:tr[docx2hub:is-tableheader-row(.)])"/>
      <xsl:variable name="next-non-vmerged-tr" as="element(w:tr)?"
        select="../following-sibling::w:tr[not(docx2hub:is-blind-vmerged-row(.))][1]"/>
      <xsl:variable name="counts" as="xs:integer*">
        <xsl:choose>
          <xsl:when test="$is-thead-tr = true() and not(../following-sibling::w:tr[1][docx2hub:is-tableheader-row(.)])">
            <xsl:sequence select="999"/>
          </xsl:when>
          <xsl:when test="$next-non-vmerged-tr/w:tc[tr:colcount(1, .) = tr:colcount(1, current())]
                                                   [docx2hub:is-blind-vmerged-cell(.)]">
            <xsl:for-each-group select="../following-sibling::w:tr[not(docx2hub:is-blind-vmerged-row(.))][. is $next-non-vmerged-tr or . &gt;&gt; $next-non-vmerged-tr]/w:tc[tr:colcount(1, .) = tr:colcount(1, current())]" 
              group-adjacent="docx2hub:is-blind-vmerged-cell(.)">
              <xsl:sequence select="count(current-group())"/>
            </xsl:for-each-group>
          </xsl:when>
          <xsl:otherwise>
            <xsl:sequence select="999"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="$counts[1] != 999">
        <xsl:attribute name="morerows" select="$counts[1]"/>
      </xsl:if>
      <!-- uncommented, mk: do not invent information where none exist -->
      <!--<xsl:attribute name="css:border-bottom-style" select="'none'" />-->
    </xsl:if>
  </xsl:template>

  <!-- recursively count preceding columns, including spans -->
  <xsl:function name="tr:colcount" as="xs:double">
    <xsl:param name="count" as="xs:double"/>
    <xsl:param name="cell"  as="element(*)"/>
    
    <xsl:choose>
      <xsl:when test="$cell/preceding-sibling::w:tc">
        <xsl:variable name="span" select="if ($cell/preceding-sibling::w:tc[1]/w:tcPr/w:gridSpan/@w:val) 
                                          then $cell/preceding-sibling::w:tc[1]/w:tcPr/w:gridSpan/@w:val
                                          else 1"/>
        <xsl:sequence select="tr:colcount($count + $span, $cell/preceding-sibling::w:tc[1])" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$count"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <!-- Tabellen innerhalb einer Tabellenzelle -->
  <xsl:template match="w:tbl" mode="tables">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>


  <!-- handle w:bidiVisual: mirror columns -->
  <xsl:template match="w:tbl[w:tblPr[@docx2hub:generated-bidiVisual]]/w:tblGrid" mode="docx2hub:remove-redundant-run-atts">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each select="w:gridCol">
        <xsl:sort select="position()" order="descending"/>
        <xsl:apply-templates select="." mode="#current"/>
      </xsl:for-each>
      <xsl:apply-templates select="* except w:gridCol" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <!-- handle w:bidiVisual: mirror columns -->
  <xsl:template match="w:tbl[w:tblPr[@docx2hub:generated-bidiVisual]]/w:tr" mode="docx2hub:remove-redundant-run-atts">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each select="*">
        <xsl:sort select="position()" order="descending"/>
        <xsl:apply-templates select="." mode="#current"/>
      </xsl:for-each>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>
