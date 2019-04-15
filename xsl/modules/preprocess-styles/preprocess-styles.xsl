<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  exclude-result-prefixes="xsl xs docx2hub"
  version="2.0">
  
  <xsl:variable name="word-2013-tablestyle-rules" select="
    every $b in (//w:settings/w:compat//(
      (xs:decimal(w:compatSetting[@w:name = 'compatibilityMode']/@w:val), 15)[1] lt 15
      and
      not(w:compatSetting[@w:name = 'differentiateMultirowTableHeaders'][@w:val = (1, 'true')])
    )) satisfies not($b)" as="xs:boolean"/>
  
  <xsl:template match="w:styles/w:style[@w:type = 'table']" exclude-result-prefixes="#all" mode="docx2hub:preprocess-styles">
    <!-- (w:style)* -->
    <xsl:variable name="self" select="."/>
    <xsl:variable name="doc-styles" select=".."/>
    <xsl:for-each select="('', w:tblStylePr/@w:type)">
      <xsl:variable name="stripped-style" as="node()">
        <w:style>
          <xsl:sequence select="$self/@* except $self/@w:styleId"/>
          <xsl:attribute name="w:styleId" select="($self[current() = '']/@w:styleId, concat($self/@w:styleId, '-', .))[1]"/>
          <xsl:variable name="bare-props" select="$self/node()[not(self::w:tblStylePr)]" as="node()*"/>
          <xsl:sequence
            select="docx2hub:consolidate-style-properties($self/element()[self::w:tblStylePr[@w:type = current()]]/element(), $bare-props)"
          />
        </w:style>
      </xsl:variable>
      <xsl:element name="{$self/name()}">
        <xsl:sequence
          select="$stripped-style/@*, docx2hub:get-style-properties($stripped-style, $doc-styles)"
        />
      </xsl:element>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template match="w:tc[empty(w:tcPr)]" exclude-result-prefixes="#all" mode="docx2hub:preprocess-styles">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:element name="w:tcPr"/>
      <xsl:apply-templates select="node()"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:function name="docx2hub:get-style-properties" as="node()*">
    <xsl:param name="style" as="node()"/><!-- w:style -->
    <xsl:param name="styles" as="node()*"/><!-- (w:style)* -->
    <xsl:variable name="based-on" select="$style/w:basedOn/@w:val"/>
    <xsl:variable name="resolved-parent-style"
      select="if (exists($styles/node()[@w:styleId = $based-on]))
        then docx2hub:get-style-properties($styles/node()[@w:styleId = $based-on], $styles)
        else ()"
      as="node()*"/>
    <xsl:sequence select="docx2hub:consolidate-style-properties($style/element(), $resolved-parent-style)"/>
  </xsl:function>

  <xsl:function name="docx2hub:consolidate-style-properties" as="node()*">
    <xsl:param name="style1" as="node()*"/>
    <xsl:param name="style2" as="node()*"/>
    <xsl:for-each select="$style1">
      <xsl:variable name="ptype" select="concat(local-name(), @w:type)"/>
      <xsl:choose>
        <xsl:when test="self::w:tblBorders and not($word-2013-tablestyle-rules)">
          <xsl:sequence select="."/>
        </xsl:when>
        <xsl:when test="not(node()) and @*">
          <!-- copy flat prop from style1 -->
          <xsl:element name="{name()}">
            <xsl:sequence select="$style2[concat(local-name(), @w:type) = $ptype]/@*[docx2hub:is-inheritable-attribute(.)], @*"/>
          </xsl:element>
        </xsl:when>
        <xsl:otherwise>
          <!-- recursive for nested props -->
          <xsl:element name="{name()}">
            <xsl:sequence select="@*"/>
            <xsl:sequence
              select="docx2hub:consolidate-style-properties(element(), $style2[concat(local-name(), @w:type) = $ptype]/element())"
            />
          </xsl:element>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
    <xsl:for-each select="$style2">
      <xsl:variable name="ptype" as="xs:string"
        select="concat('', local-name(), @w:type[count($style1/element()[local-name() = current()/local-name()]) gt 1])"/>
      <xsl:choose>
        <xsl:when test="$style1[concat('', local-name(), @w:type[count($style1/element()[local-name() = current()/local-name()]) gt 1]) = $ptype]">
          <!-- this one was copied from style1 already -->
        </xsl:when>
        <xsl:otherwise>
          <!-- props added by style2 -->
          <xsl:sequence select="."/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:function>
  
  <xsl:function name="docx2hub:is-inheritable-attribute" as="xs:boolean">
    <xsl:param name="attribute" as="attribute()"/>
    <xsl:sequence select="not($attribute/local-name() = ('hanging'))"/>
  </xsl:function>

  <xsl:template match="w:tbl/w:tr" mode="docx2hub:resolve-tblBorders">
    <xsl:next-match>
      <xsl:with-param name="tr-pos"
        select="count(preceding-sibling::*) + 1 - count(preceding-sibling::*[not(self::w:tr)])" tunnel="yes"/>
    </xsl:next-match>
  </xsl:template>
  
  <xsl:template match="w:tbl/w:tr/w:tc" mode="docx2hub:resolve-tblBorders">
    <xsl:param name="tr-pos" tunnel="yes"/>
    <xsl:variable name="self" select="."/>
    <xsl:variable name="tc-pos" select="count(preceding-sibling::*) + 1 - count(preceding-sibling::*[not(self::w:tc)])"/>
    <xsl:variable name="pos" select="$tc-pos, $tr-pos, ../count(w:tc), ../../count(w:tr)" as="xs:decimal+">
      <!-- values are: tc-pos, tr-pos, last-tc, last-tr -->
    </xsl:variable>
    <xsl:variable name="style-name" as="xs:string*">
      <xsl:variable name="prefix" select="../../w:tblPr/w:tblStyle/@w:val"/>
      <xsl:variable name="lk" as="xs:boolean+"
      select="
      ((../../w:tblPr/w:tblLook, ../w:tblPrEx/w:tblBorders)/@w:firstRow)[1] = (1, 'true'),
      ((../../w:tblPr/w:tblLook, ../w:tblPrEx/w:tblBorders)/@w:firstColumn)[1] = (1, 'true'),
      ((../../w:tblPr/w:tblLook, ../w:tblPrEx/w:tblBorders)/@w:lastRow)[1] = (1, 'true'),
      ((../../w:tblPr/w:tblLook, ../w:tblPrEx/w:tblBorders)/@w:lastColumn)[1] = (1, 'true')"/>
      <xsl:variable name="suffix" as="xs:string*">
        <xsl:variable name="possible-suffixes" as="xs:string*">
          <xsl:choose>
            <xsl:when test="$lk[2] and $lk[1] and $pos[1] eq 1 and $pos[2] eq 1">
              <xsl:sequence select="'nwCell', 'firstRow', 'firstCol'"/>
            </xsl:when>
            <xsl:when test="$lk[2] and $lk[3] and $pos[1] eq 1 and $pos[2] eq $pos[4]">
              <xsl:sequence select="'swCell', 'lastRow', 'firstCol'"/>
            </xsl:when>
            <xsl:when test="$lk[4] and $lk[1] and $pos[1] eq $pos[3] and $pos[2] eq 1">
              <xsl:sequence select="'neCell', 'firstRow', 'lastCol'"/>
            </xsl:when>
            <xsl:when test="$lk[4] and $lk[3] and $pos[1] eq $pos[3] and $pos[2] eq $pos[4]">
              <xsl:sequence select="'seCell', 'lastRow', 'lastCol'"/>
            </xsl:when>
            <xsl:when test="$lk[1] and ($pos[2] eq 1 or count((../preceding-sibling::w:tr/w:trPr/w:tblHeader, ../w:trPr/w:tblHeader)) eq $tr-pos)">
              <xsl:sequence select="'firstRow'"/>
            </xsl:when>
            <xsl:when test="$lk[2] and $pos[1] eq 1">
              <xsl:sequence select="'firstCol'"/>
            </xsl:when>
            <xsl:when test="$lk[3] and $pos[2] eq $pos[4]">
              <xsl:sequence select="'lastRow'"/>
            </xsl:when>
            <xsl:when test="$lk[4] and $pos[1] eq $pos[3]">
              <xsl:sequence select="'lastCol'"/>
            </xsl:when>
            <!-- TODO: H/V-banding props -->
            <xsl:otherwise/>
          </xsl:choose>
        </xsl:variable>
        <xsl:sequence select="for $s in $possible-suffixes return $s[exists($self/ancestor::*[last()]//w:style[@w:styleId = concat($prefix, '-', $s)])]"/>
      </xsl:variable>
      <xsl:sequence select="$prefix[empty($suffix)] ,for $s in $suffix return string-join(($prefix, ('-', $s)[not($s= '')]), '')"/>
    </xsl:variable>
    <xsl:variable name="sty" select="for $s in $style-name return $self/ancestor::*[last()]//w:style[@w:styleId = concat('', $s)]" as="element()*"/>
    <xsl:next-match>
      <xsl:with-param name="tc-pos" select="$tc-pos" tunnel="yes"/>
      <xsl:with-param name="tblsty" select="$sty" tunnel="yes"/>
    </xsl:next-match>
  </xsl:template>

  <xsl:template match="w:tbl[empty(.//(w:tblPrEx | w:tblPr)/w:tblCellSpacing[matches(@w:w, '[1-9]')])]/w:tblPr"
    mode="docx2hub:resolve-tblBorders">
    <!-- if tblCellSpacing is 0, its sufficient to apply borders on cells only -->
    <xsl:copy>
      <xsl:apply-templates select="@*, node() except w:tblBorders" mode="#current"/>
      <xsl:variable name="tblBorder" select="w:tblBorders/*" as="node()*"/>
      <w:tblBorders>
        <xsl:for-each select="'top', 'left', 'bottom', 'right'">
          <xsl:element name="w:{.}">
            <xsl:attribute name="w:val" select="'nil'"/>
          </xsl:element>
        </xsl:for-each>
      </w:tblBorders>
    </xsl:copy>
  </xsl:template>

  <xsl:template mode="docx2hub:resolve-tblBorders"
    match="w:tbl/w:tr/w:tc/w:tcPr[empty(ancestor::*/(w:tblPrEx|w:tblPr)/w:tblCellSpacing[matches(@w:w, '[1-9]')])]">
    <xsl:param name="tr-pos" tunnel="yes"/>
    <xsl:param name="tc-pos" tunnel="yes"/>
    <xsl:param name="tblsty" tunnel="yes" as="element()*"/>
    <xsl:variable name="self" select="."/>
    <xsl:variable name="pos" select="$tc-pos, $tr-pos, count(../../w:tc), ../../../count(w:tr)" as="xs:decimal+">
      <!-- values are: tc-pos, tr-pos, last-tc, last-tr -->
    </xsl:variable>
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates select="node() except w:tcBorders" mode="#current"/>
      <xsl:element name="w:tcBorders">
        <xsl:for-each select="'top','left','bottom','right'">
          <xsl:variable name="outer-border" select="
            (. = 'top' and $pos[2] eq 1) or
            (. = 'left' and $pos[1] eq 1) or
            (. = 'bottom' and $pos[2] eq $pos[4]) or
            (. = 'right' and $pos[1] eq $pos[3])" as="xs:boolean"/>
          <xsl:variable name="tblBorder-by-style" as="element()?">
            <xsl:variable name="borders"
              select="$tblsty/w:tblPr/w:tblBorders, $tblsty/w:tcPr/w:tcBorders"/>
            <xsl:element name="w:{.}">
              <xsl:choose>
                <xsl:when test="$outer-border">
                  <xsl:sequence
                    select="($borders/w:*[local-name() = current()])[last()]/@*"
                  />
                </xsl:when>
                <!-- inner tcBorder -->
                <xsl:otherwise>
                  <xsl:sequence
                    select="($borders/w:insideH[current() = ('top', 'bottom')], $borders/w:insideV[current() = ('left', 'right')])[1]/@*"
                  />
                </xsl:otherwise>
              </xsl:choose>
            </xsl:element>
          </xsl:variable>
          <xsl:variable name="ancestors" as="element()*">
            <xsl:variable name="resolve-inside" as="element()*">
              <xsl:choose>
                <xsl:when test="$outer-border">
                  <xsl:sequence select="$self/../../w:tblPrEx/w:tblBorders, $self/../../../w:tblPr/w:tblBorders"/>
                </xsl:when>
                <!-- inner tcBorder -->
                <xsl:otherwise>
                  <xsl:element name="w:tblBorders">
                    <xsl:element name="w:{.}">
                      <xsl:sequence
                        select="(($self/../../w:tblPrEx/w:tblBorders, $self/../../../w:tblPr/w:tblBorders)/(w:insideH[current() = ('top', 'bottom')], w:insideV[current() = ('left', 'right')]))[1]/@*"
                      />
                    </xsl:element>
                  </xsl:element>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:sequence select="$resolve-inside[.//@*], $tblsty/w:tcPr/w:tcBorders"/>
          </xsl:variable>
          <xsl:variable name="adhoc-border"
            select="(
              $self/w:tcBorders/w:*[local-name() = current()],
              $ancestors/w:*[local-name() = current()]
            )[1]"
            as="element()?"
          />
          <xsl:variable name="border">
            <xsl:choose>
              <xsl:when test="exists($adhoc-border)">
                <xsl:sequence select="$adhoc-border"/>
              </xsl:when>
              <xsl:when test="exists($tblBorder-by-style/@*)">
                <xsl:sequence select="$tblBorder-by-style"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:element name="w:{.}">
                  <xsl:attribute name="w:val" select="'nil'"/>
                </xsl:element>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:apply-templates select="$border" mode="#current"/>
        </xsl:for-each>
      </xsl:element>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:p/w:r" mode="docx2hub:resolve-tblBorders">
    <xsl:param name="tblsty" tunnel="yes" as="element()*"/>
    <xsl:variable name="self" select="."/>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:variable name="consolidated-props" as="element()?">
        <xsl:variable name="para-style" select="//w:style[@w:styleId = current()/../w:pPr/w:pStyle/@w:val]/w:rPr"
          as="element()*"/>
        <xsl:variable name="run-style" select="//w:style[@w:styleId = current()/w:rPr/w:rStyle/@w:val]/w:rPr"
          as="element()*"/>
        <!-- all toggle properties according to §17.7.3: 
          'b', 'bCs', 'caps', 'emboss', 'i', 'iCs', 'imprint', 'outline', 'shadow', 'smallCaps', 'strike', 'vanish' -->
        <xsl:variable name="style-toggles" as="element()*">
          <xsl:variable name="table-toggles">
            <xsl:for-each-group select="$tblsty//(w:b, w:bCs, w:caps, w:emboss, w:i, w:iCs, w:imprint, w:outline, w:shadow, w:smallCaps, w:strike, w:vanish)" group-by="name()">
              <xsl:sequence select="current-group()[last()]"/>
            </xsl:for-each-group>
          </xsl:variable>
          <xsl:sequence select="($table-toggles, $para-style, $run-style)//(w:b, w:bCs, w:caps, w:emboss, w:i, w:iCs, w:imprint, w:outline, w:shadow, w:smallCaps, w:strike, w:vanish)"/>
        </xsl:variable>
        <xsl:if test="exists((w:rPr, $style-toggles))">
          <xsl:element name="w:rPr">
            <xsl:apply-templates select="w:rPr/node()"/>
            <xsl:for-each-group select="$style-toggles" group-by="name()">
              <xsl:variable name="name" select="current-group()[1]/name()"/>
              <xsl:variable name="toggle-prop"
                select="count(($style-toggles[name() = $name][docx2hub:is-toggled(.)]))"
                as="xs:decimal"/>
              <xsl:if test="not($self/w:rPr/*[name() = $name]) and $toggle-prop gt 0">
                <xsl:element name="{$name}">
                  <xsl:attribute name="w:val"
                    select="($toggle-prop mod 2)[1]"
                  />
                </xsl:element>
              </xsl:if>
            </xsl:for-each-group>
          </xsl:element>
        </xsl:if>
      </xsl:variable>
      <xsl:apply-templates select="$consolidated-props, node() except w:rPr"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:function name="docx2hub:is-toggled" as="xs:boolean">
    <xsl:param name="prop" as="element()"/>
    <xsl:sequence select="empty($prop/@w:val) or $prop/@w:val = ('1', 'true')"/>
  </xsl:function>
  
  <xsl:template match="w:r[count(child::*)=1][w:rPr]" mode="docx2hub:preprocess-styles"/>
  
</xsl:stylesheet>
