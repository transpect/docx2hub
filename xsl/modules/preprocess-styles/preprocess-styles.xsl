<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  exclude-result-prefixes="xsl xs docx2hub"
  version="2.0">
  
  <xsl:variable name="word-2013-tablestyle-rules" select="
    every $b in (//w:settings/w:compat//(
      (xs:decimal(w:compatSetting[@w:name = 'compatibilityMode']/@w:val), 15)[1] lt 15
      and
      not(w:compatSetting[@w:name = 'differentiateMultirowTableHeaders'][@w:val = (1, 'true')])
    )) satisfies not($b)" as="xs:boolean"/>

  <xsl:template match="w:tbl" mode="docx2hub:resolve-tblBorders">
    <xsl:variable name="based-on-chain" as="node()*"
      select="reverse(
          docx2hub:based-on-chain(
            //w:style[@w:styleId = current()/w:tblPr/w:tblStyle/@w:val]
          )/node()
        )"/>
    <xsl:variable name="first-border-style" select="($based-on-chain[.//w:tblBorders])[last()]"/>
    <xsl:next-match>
      <xsl:with-param name="tblSty" as="node()*" tunnel="yes"
        select="if ($word-2013-tablestyle-rules) then $based-on-chain else $first-border-style"/>
    </xsl:next-match>
  </xsl:template>

  <xsl:template match="w:tbl/w:tr" mode="docx2hub:resolve-tblBorders">
    <xsl:next-match>
      <xsl:with-param name="tr-pos" tunnel="yes"
        select="count(preceding-sibling::*) + 1 - count(preceding-sibling::*[not(self::w:tr)])"/>
    </xsl:next-match>
  </xsl:template>
  
  <xsl:template match="w:tbl/w:tr/w:tc" mode="docx2hub:resolve-tblBorders">
    <xsl:param name="tr-pos" tunnel="yes"/>
    <xsl:param name="tblSty" as="node()*" tunnel="yes"/>
    <xsl:variable name="self" select="."/>
    <xsl:variable name="tc-pos" select="count(preceding-sibling::*) + 1 - count(preceding-sibling::*[not(self::w:tc)])"/>
    <xsl:variable name="pos" select="$tc-pos, $tr-pos, ../count(w:tc), ../../count(w:tr)" as="xs:decimal+">
      <!-- values are: tc-pos, tr-pos, last-tc, last-tr -->
    </xsl:variable>
    <xsl:next-match>
      <xsl:with-param name="tc-pos" select="$tc-pos" tunnel="yes"/>
      <xsl:with-param name="tblsty" select="$tblSty" tunnel="yes"/>
      <xsl:with-param name="tblStylePr-name" tunnel="yes"
                      select="docx2hub:active-cnf-style-name((w:tcPr/w:cnfStyle,parent::w:tr/w:trPr/w:cnfStyle))"/>
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
    <xsl:param name="tblsty" tunnel="yes" as="node()*"/>
    <xsl:param name="tblStylePr-name" tunnel="yes" as="xs:string*"/>
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
            <xsl:variable name="borders" as="node()*">
              <xsl:variable name="pre-borders" as="node()*"
                select="$tblsty/w:tblPr/w:tblBorders,
                $tblsty/w:tblStylePr[@w:type = $tblStylePr-name]/w:tcPr/w:tcBorders,
                $tblsty/w:tcPr/w:tcBorders"/>
              <xsl:sequence select="$pre-borders except $pre-borders[$word-2013-tablestyle-rules][self::w:tblBorder[following-sibling::w:tblBorder]]"/>
            </xsl:variable>
            <xsl:element name="w:{.}">
              <xsl:choose>
                <xsl:when test="$outer-border">
                  <xsl:sequence
                    select="($borders/w:*[local-name() = current()])[last()]/@*"
                  />
                </xsl:when>
                <!-- inner tcBorder -->
                <xsl:when test="$tblsty/w:tblStylePr[@w:type = $tblStylePr-name]/w:tcPr/w:tcBorders/w:*[local-name() = current()]">
                  <xsl:sequence
                    select="($tblsty/w:tblStylePr[@w:type = $tblStylePr-name]/w:tcPr/w:tcBorders/w:*[local-name() = current()])[last()]/@*"
                  />
                </xsl:when>
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
  
  <xsl:function name="docx2hub:active-cnf-style-name" as="xs:string*">
    <xsl:param name="cnfStyle" as="element(w:cnfStyle)*"/>
    <xsl:choose>
      <xsl:when test="$cnfStyle/@w:firstRowFirstColumn = ('1', 'yes')">
        <xsl:sequence select="'nwCell'"/>
      </xsl:when>
      <xsl:when test="$cnfStyle/@w:lastRowFirstColumn = ('1', 'yes')">
        <xsl:sequence select="'swCell'"/>
      </xsl:when>
      <xsl:when test="$cnfStyle/@w:firstRowLastColumn = ('1', 'yes')">
        <xsl:sequence select="'neCell'"/>
      </xsl:when>
      <xsl:when test="$cnfStyle/@w:lastRowLastColumn = ('1', 'yes')">
        <xsl:sequence select="'seCell'"/>
      </xsl:when>
      <!-- TODO: H/V-banding props -->
      <xsl:otherwise>
        <xsl:for-each select="
          ($cnfStyle/(
          @w:firstRow, @w:FirstColumn,
          @w:lastRow, @w:LastColumn
          )[. = ('1', 'yes')])[1]">
          <xsl:sequence select="replace(local-name(), 'umn$', '')"/>
          <!-- 'firstColumn' to 'firstCol' -->
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <!-- Collateral: Map built-in style names in localized files. Dealing only with headings for the time being. 
    There is a list of built-in styles at https://bettersolutions.com/word/styles/list-all-built-in-styles.htm -->
  
  <xsl:template match="w:style[@w:type='paragraph'][not(@w:customStyle='1')]/@w:styleId[../w:name/@w:val[matches(., '^heading \d$')]]
                                                              [not(matches(., '^Heading\d'))]" 
    mode="docx2hub:resolve-tblBorders">
    <xsl:attribute name="{name()}" select="concat('Heading', replace(../w:name/@w:val, 'heading ', ''))"/>
  </xsl:template>
  
  <xsl:template match="w:style[@w:type='paragraph']
                              [@w:customStyle='1']/@w:styleId[matches(., '^Heading\d')]
                              [some $p in (/w:root/w:styles/w:style[@w:type='paragraph'][not(@w:customStyle='1')]/w:name/@w:val) 
                               satisfies . = concat('Heading', replace($p, 'heading ', ''))]" 
                mode="docx2hub:resolve-tblBorders">
    <xsl:variable name="count" select="for $s in /w:root/w:styles/w:style[@w:type='paragraph']/@w:styleId[matches(.,concat('^',current(),'\d+$'))]
                                       return xs:integer(replace($s, ., ''))" as="xs:integer *"/>
    <xsl:attribute name="{name()}" select="concat(., if (empty($count)) then '0' else max($count)+1)"/>
  </xsl:template>
  
  <xsl:template match="*[name() = ('w:pStyle', 'w:basedOn')]
    /@w:val[. = /w:root/w:styles/w:style[@w:type='paragraph']
                                        [matches(w:name/@w:val, '^heading \d$') or matches(@w:styleId,'^Heading\d+$')]/@w:styleId]"
                mode="docx2hub:resolve-tblBorders">
    <xsl:attribute name="{name()}">
      <xsl:apply-templates select="/w:root/w:styles/w:style[@w:type='paragraph'][@w:styleId = current()]/@w:styleId"
        mode="#current"/>
    </xsl:attribute> 
  </xsl:template>
  
  <xsl:template match="w:p[w:pPr/w:rPr]/m:oMathPara" mode="docx2hub:resolve-tblBorders">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates select="parent::w:p/w:pPr/w:rPr" mode="#current">
        <xsl:with-param name="in-math" select="true()"/>
      </xsl:apply-templates>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:p[m:oMathPara]/w:pPr/w:rPr" mode="docx2hub:resolve-tblBorders">
    <xsl:param name="in-math" select="false()"/>
    <xsl:choose>
      <xsl:when test="$in-math">
        <xsl:copy>
          <xsl:apply-templates select="@*" mode="#current"/>
          <xsl:apply-templates select="w:b | w:i | w:color | w:sz | w:caps" mode="#current"/>
        </xsl:copy>        
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>
