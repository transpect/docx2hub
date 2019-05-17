<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl	= "http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:tr="http://transpect.io"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:v="urn:schemas-microsoft-com:vml" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:word200x="http://schemas.microsoft.com/office/word/2003/wordml"
  xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:extendedProps="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:customProps="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0"
  exclude-result-prefixes="w dbk r rel tr m v wp">

  <xsl:import href="propmap.xsl"/>
  <xsl:import href="http://transpect.io/xslt-util/colors/xsl/colors.xsl"/>

  <xsl:key name="docx2hub:style" match="w:style" use="@w:styleId" />
  <xsl:key name="docx2hub:style-by-role" match="css:rule | dbk:style" use="if ($hub-version eq '1.0') then @role else @name" />
  <xsl:key name="docx2hub:font-by-name" match="w:font[@w:name]" use="@w:name"/>
  <xsl:key name="docx2hub:header-footer-ref-by-id" match="w:headerReference|w:footerReference" use="@r:id"/>
  
  <xsl:template match="/" mode="docx2hub:add-props">
    <xsl:apply-templates select="w:root/w:document/w:body" mode="#current" />
  </xsl:template>
  
  <xsl:template match="@srcpath" mode="docx2hub:add-props">
    <xsl:attribute name="srcpath" select="substring-after(., escape-html-uri(/w:root/@xml:base))" />
  </xsl:template>

  <xsl:template match="w:body" mode="docx2hub:add-props">
    <xsl:element name="{if ($hub-version eq '1.0') then 'Body' else 'hub'}">
      <xsl:namespace name="a">http://schemas.openxmlformats.org/drawingml/2006/main</xsl:namespace>
      <xsl:namespace name="cp">http://schemas.openxmlformats.org/package/2006/metadata/core-properties</xsl:namespace>
      <xsl:namespace name="ct">http://schemas.openxmlformats.org/package/2006/content-types</xsl:namespace>
      <xsl:namespace name="cx">http://schemas.microsoft.com/office/drawing/2014/chartex</xsl:namespace>
      <xsl:namespace name="cx1">http://schemas.microsoft.com/office/drawing/2015/9/8/chartex</xsl:namespace>
      <xsl:namespace name="customProps">http://schemas.openxmlformats.org/officeDocument/2006/custom-properties</xsl:namespace>
      <xsl:namespace name="docx2hub">http://transpect.io/docx2hub</xsl:namespace>
      <xsl:namespace name="extendedProps">http://schemas.openxmlformats.org/officeDocument/2006/extended-properties</xsl:namespace>
      <xsl:namespace name="m">http://schemas.openxmlformats.org/officeDocument/2006/math</xsl:namespace>
      <xsl:namespace name="mc">http://schemas.openxmlformats.org/markup-compatibility/2006</xsl:namespace>
      <xsl:namespace name="mml">http://www.w3.org/1998/Math/MathML</xsl:namespace>
      <xsl:namespace name="o">urn:schemas-microsoft-com:office:office</xsl:namespace>
      <xsl:namespace name="pkg">http://schemas.microsoft.com/office/2006/xmlPackage</xsl:namespace>
      <xsl:namespace name="r">http://schemas.openxmlformats.org/officeDocument/2006/relationships</xsl:namespace>
      <xsl:namespace name="rel">http://schemas.openxmlformats.org/package/2006/relationships</xsl:namespace>
      <xsl:namespace name="tr">http://transpect.io</xsl:namespace>
      <xsl:namespace name="v">urn:schemas-microsoft-com:vml</xsl:namespace>
      <xsl:namespace name="ve">http://schemas.openxmlformats.org/markup-compatibility/2006</xsl:namespace>
      <xsl:namespace name="w">http://schemas.openxmlformats.org/wordprocessingml/2006/main</xsl:namespace>
      <xsl:namespace name="w10">urn:schemas-microsoft-com:office:word</xsl:namespace>
      <xsl:namespace name="w14">http://schemas.microsoft.com/office/word/2010/wordml</xsl:namespace>
      <xsl:namespace name="w15">http://schemas.microsoft.com/office/word/2012/wordml</xsl:namespace>
      <xsl:namespace name="w16cid">http://schemas.microsoft.com/office/word/2016/wordml/cid</xsl:namespace>
      <xsl:namespace name="w16se">http://schemas.microsoft.com/office/word/2015/wordml/symex</xsl:namespace>
      <xsl:namespace name="wne">http://schemas.microsoft.com/office/word/2006/wordml</xsl:namespace>
      <xsl:namespace name="word200x">http://schemas.microsoft.com/office/word/2003/wordml</xsl:namespace>
      <xsl:namespace name="wp">http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing</xsl:namespace>
      <xsl:namespace name="wp14">http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing</xsl:namespace>
      <xsl:namespace name="wpc">http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas</xsl:namespace>
      <xsl:namespace name="wpg">http://schemas.microsoft.com/office/word/2010/wordprocessingGroup</xsl:namespace>
      <xsl:namespace name="wpi">http://schemas.microsoft.com/office/word/2010/wordprocessingInk</xsl:namespace>
      <xsl:namespace name="wps">http://schemas.microsoft.com/office/word/2010/wordprocessingShape</xsl:namespace>
      <xsl:namespace name="wx">http://schemas.microsoft.com/office/word/2003/auxHint</xsl:namespace>
      <xsl:namespace name="xlink">http://www.w3.org/1999/xlink</xsl:namespace>
      <xsl:namespace name="xs">http://www.w3.org/2001/XMLSchema</xsl:namespace>
      <xsl:if test="../../w:styles/@xml:lang">
        <!-- might be superseded by the most frequently used language in a later mode --> 
        <xsl:attribute name="xml:lang" select="replace(../../w:styles/@xml:lang, '\-[A-Z]+$', '')"/>  
      </xsl:if>
      <xsl:attribute name="version" select="concat('5.1-variant le-tex_Hub-', $hub-version)"/>
      <xsl:attribute name="css:version" select="concat('3.0-variant le-tex_Hub-', $hub-version)" />
      <xsl:if test="not($hub-version eq '1.0')">
        <xsl:attribute name="css:rule-selection-attribute" select="'role'" />
      </xsl:if>
      <info>
        <keywordset role="hub">
          <keyword role="formatting-deviations-only">true</keyword>
          <keyword role="source-type">docx</keyword>
          <xsl:if test="/w:root/@xml:base != ''">
            <keyword role="source-dir-uri">
              <xsl:value-of select="/w:root/@xml:base" />
            </keyword>
            <keyword role="archive-dir-uri">
              <xsl:value-of select="concat(replace(/w:root/@xml:base, '^(.+)/[^/]+/?', '$1'), '/')"/>
            </keyword>
            <keyword role="archive-uri">
              <xsl:value-of select="/w:root/@archive-uri" />
            </keyword>
            <keyword role="archive-uri-local">
              <xsl:value-of select="/w:root/@archive-uri-local" />
            </keyword>
          </xsl:if>
          <keyword role="source-basename">
            <!-- /w:root/@xml:base example: file:/data/docx/M_001.docx.tmp/ -->
            <xsl:value-of select="replace(/w:root/@xml:base, '^.*/(.+)\.doc[xm](\.tmp/?)?$', '$1')"/>
          </keyword>
          <xsl:if test="/w:root/w:containerProps/extendedProps:Properties/extendedProps:Application">
            <keyword role="source-application">
              <xsl:value-of select="/w:root/w:containerProps/extendedProps:Properties/extendedProps:Application"/>
            </keyword>
          </xsl:if>
          <xsl:if test="not($mathtype2mml eq 'no')">
            <keyword role="mathtype2mml">
              <xsl:value-of select="if ($mathtype2mml = 'yes') then 'true' else $mathtype2mml"/>
            </keyword>
          </xsl:if>
        </keywordset>
        <xsl:if test="exists(../../w:containerProps/(extendedProps:Properties|cp:coreProperties))">
          <keywordset role="docProps">
            <xsl:for-each select="../../w:containerProps/cp:*/*[not(*)]">
              <keyword role="{name()}">
                <xsl:value-of select="."/>
              </keyword>
            </xsl:for-each>
            <xsl:for-each select="../../w:containerProps/extendedProps:Properties/extendedProps:*[not(*)]">
              <keyword role="extendedProps:{local-name()}">
                <xsl:value-of select="."/>
              </keyword>
            </xsl:for-each>
          </keywordset>
        </xsl:if>
        <xsl:if test="exists(../../w:settings/w:docVars/w:docVar)">
          <keywordset role="docVars">
            <xsl:for-each select="../../w:settings/w:docVars/w:docVar">
              <keyword role="{@w:name}">
                <xsl:value-of select="@w:val"/>
              </keyword>
            </xsl:for-each>
          </keywordset>
        </xsl:if>
        <xsl:if test="exists(/w:root/w:containerProps/customProps:Properties/customProps:property)">
          <keywordset role="custom-meta">
            <xsl:apply-templates mode="#current" 
              select="/w:root/w:containerProps/customProps:Properties/customProps:property"/>
          </keywordset>
        </xsl:if>
        <xsl:if test="$field-vars='yes'">
          <keywordset role="fieldVars"/>
        </xsl:if>
        <xsl:choose>
          <xsl:when test="$hub-version eq '1.0'">
            <xsl:call-template name="docx2hub:hub-1.0-styles">
              <xsl:with-param name="version" select="$hub-version" tunnel="yes"/>
              <xsl:with-param name="contexts" select="., 
                                                      /w:root/w:footnotes,
                                                      /w:root/w:endnotes"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="docx2hub:hub-1.1-styles">
              <xsl:with-param name="version" select="$hub-version" tunnel="yes"/>
              <xsl:with-param name="contexts" select="., 
                                                      /w:root/w:header,
                                                      /w:root/w:footer,
                                                      /w:root/w:footnotes, 
                                                      /w:root/w:endnotes, 
                                                      /w:root/w:footer/w:ftr[$convert-footer]"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </info>
      <xsl:if test="$include-header-and-footer eq 'yes'">
        <xsl:for-each-group select="/w:root/w:header/w:hdr|/w:root/w:footer/w:ftr" group-by="parent::*/local-name()">
          <div role="docx2hub:{current-grouping-key()}-spec">
            <xsl:for-each select="current-group()">
              <xsl:variable name="header-footer-basename" as="xs:string"
                            select="replace(@xml:base, '^.+/(.+)$', '$1')"/>
              <xsl:variable name="header-footer-id" as="attribute(Id)"
                            select="/w:root/w:docRels/rel:Relationships/rel:Relationship
                                    [@Type = ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
                                              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
                                     and @Target eq $header-footer-basename]/@Id"/>
              <div role="docx2hub:{current-grouping-key()}" 
                   condition="{key('docx2hub:header-footer-ref-by-id', $header-footer-id)/@w:type}">
                <xsl:apply-templates mode="#current"/>
              </div>
            </xsl:for-each>        
          </div>
        </xsl:for-each-group>
      </xsl:if>
      <xsl:apply-templates select="../../w:numbering" mode="#current"/>
      <xsl:sequence select="../../w:docRels, ../../w:footnoteRels, ../../w:endnoteRels, ../../w:commentRels, ../../w:fonts"/>
      <xsl:apply-templates select="../../w:comments, ../../w:footnotes, ../../w:endnotes" mode="#current"/>
      <xsl:apply-templates mode="#current"/>
      <xsl:apply-templates select="../../w:footer/w:ftr[$convert-footer]" mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="customProps:property" mode="docx2hub:add-props">
    <keyword role="{@name}">
      <xsl:value-of select="*"/>
    </keyword>
  </xsl:template>
  
  <xsl:template match="*[ancestor-or-self::w:numbering]" mode="docx2hub:add-props" priority="-1">
    <xsl:copy>
      <xsl:sequence select="@*"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template name="docx2hub:hub-1.0-styles">
    <xsl:param name="contexts" as="element(*)+"/>
    <styles>
      <parastyles>
        <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:pStyle/@w:val))
          union
          (: Resolve linked char styles as their respective para styles :)
          key(
           'docx2hub:style',
            key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))/w:link/@w:val
          )" mode="#current">
          <xsl:sort select="@w:styleId" />
        </xsl:apply-templates>
      </parastyles>
      <inlinestyles>
        <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))[not(w:link)]" mode="#current">
          <xsl:sort select="@w:styleId" />
        </xsl:apply-templates>
      </inlinestyles>
      <tablestyles>
        <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:tblStyle/@w:val))" mode="#current">
          <xsl:sort select="@w:styleId" />
        </xsl:apply-templates>
      </tablestyles>
      <cellstyles/>
    </styles>
  </xsl:template>    

  <xsl:template name="docx2hub:hub-1.1-styles">
    <xsl:param name="contexts" as="element(*)+"/>
    <css:rules>
      <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:pStyle/@w:val))[@w:type='paragraph']
        union
        key('docx2hub:style', key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))/w:link/@w:val)[@w:type='paragraph']" 
        mode="#current">
        <xsl:sort select="@w:styleId" />
      </xsl:apply-templates>
      <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))[@w:type='character']" mode="#current">
        <xsl:sort select="@w:styleId" />
      </xsl:apply-templates>
      <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:tblStyle/@w:val))[@w:type='table']" mode="#current">
        <xsl:sort select="@w:styleId" />
      </xsl:apply-templates>
    </css:rules>
  </xsl:template>
  
    
  <xsl:template match="w:style" mode="docx2hub:add-props">
    <xsl:param name="wrap-in-style-element" select="true()" as="xs:boolean"/>
    <xsl:param name="version" as="xs:string" tunnel="yes"/>
    <xsl:variable name="atts" as="element(*)*"> <!-- docx2hub:attribute, ... -->
      <xsl:apply-templates select="if (w:basedOn/@w:val) 
                                   then key('docx2hub:style', w:basedOn/@w:val) 
                                   else ()" mode="#current">
        <xsl:with-param name="wrap-in-style-element" select="false()"/>
      </xsl:apply-templates>
      <xsl:variable name="mergeable-atts" as="element(*)*"> <!-- docx2hub:attribute, ... -->
        <xsl:apply-templates select="* except w:basedOn" mode="#current" />
      </xsl:variable>
      <xsl:for-each-group select="$mergeable-atts[self::docx2hub:attribute]" group-by="@name">
        <docx2hub:attribute name="{current-grouping-key()}">
          <xsl:value-of select="if (matches(current-grouping-key(), '^css:(margin|text-indent)'))
                                then current-group()[last()]
                                else current-group()"/>
        </docx2hub:attribute>
      </xsl:for-each-group>
      <xsl:sequence select="$mergeable-atts[self::*][not(self::docx2hub:attribute)]"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$wrap-in-style-element">
        <xsl:element name="{if ($hub-version = '1.0') then 'style' else 'css:rule'}">
          <xsl:apply-templates select="." mode="docx2hub:XML-Hubformat-add-properties_layout-type"/>
          <xsl:sequence select="$atts"/>
          <xsl:choose>
            <xsl:when test="$hub-version = '1.0'">
              <docx2hub:attribute name="role">
                <xsl:value-of select="docx2hub:css-compatible-name(@w:styleId)"/>
              </docx2hub:attribute>  
            </xsl:when>
            <xsl:otherwise>
              <docx2hub:attribute name="name">
                <xsl:value-of select="docx2hub:css-compatible-name(@w:styleId)"/>
              </docx2hub:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$atts"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:lvl" mode="docx2hub:add-props">
    <xsl:copy>
      <xsl:apply-templates mode="#current" select="@*"/>
      <xsl:apply-templates mode="#current" select="*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:basedOn/@w:val" mode="docx2hub:add-props" />

  <xsl:template match="w:style"
    mode="docx2hub:XML-Hubformat-add-properties_layout-type">
    <xsl:param name="version" tunnel="yes" as="xs:string"/>
    <xsl:if test="not($version eq '1.0')">
      <xsl:attribute name="layout-type" select="if (@w:type = 'paragraph')
        then 'para'
        else if (@w:type = 'character')
        then 'inline'
        else if (@w:type = 'table')
        then 'table'        
        else 'undefined'"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:rPr[not(ancestor::m:oMath)] | w:pPr | m:oMathParaPr" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="*" mode="#current" />
  </xsl:template>
  
  <!-- The most specific style comes first, the one that it is based on comes next, etc.
    I put it in a document node in order to avoid getting them ordered in document order 
    when I output a sequence of w:style elements. -->
  <xsl:function name="docx2hub:based-on-chain" as="document-node()">
    <xsl:param name="initial" as="element(w:style)*"/>
    <xsl:variable name="next" as="element(w:style)?" 
      select="if (exists($initial)) 
              then key('docx2hub:style', $initial[last()]/w:basedOn/@w:val, root($initial[last()]))
              else ()"/>
    <xsl:choose>
      <xsl:when test="exists($next)">
        <xsl:document>
          <xsl:sequence select="docx2hub:based-on-chain(($initial, $next))/*"/>  
        </xsl:document>
      </xsl:when>
      <xsl:otherwise>
        <xsl:document>
          <xsl:sequence select="$initial"/>  
        </xsl:document>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:template match="w:style/w:pPr[w:numPr]" mode="docx2hub:add-props" priority="3">
    <xsl:variable name="based-on-chain" as="document-node()" select="docx2hub:based-on-chain(..)"/>
    <xsl:variable name="numId-chain-item" as="element(w:style)?" select="$based-on-chain/*[w:pPr/w:numPr/w:numId][1]"/>
    <xsl:variable name="numId-chain-item-pos" as="xs:integer?" 
      select="index-of(for $item in $based-on-chain/* return generate-id($item), generate-id($numId-chain-item))"/>
    <xsl:variable name="numId" as="element(w:numId)?" select="$numId-chain-item/w:pPr/w:numPr/w:numId"/>
    <xsl:variable name="ilvl" as="xs:integer?" 
      select="(
                for $i in $based-on-chain/*[w:pPr/w:numPr/w:ilvl][1]/w:pPr/w:numPr/w:ilvl/@w:val return xs:integer($i),
                0
              )[1]"/>
    <xsl:variable name="lvl" as="element(w:lvl)?" select="key(
                                                           'abstract-numbering-by-id', 
                                                           key(
                                                             'numbering-by-id', 
                                                             $numId/@w:val 
                                                           )/w:abstractNumId/@w:val
                                                         )/w:lvl[@w:ilvl = $ilvl]"/>
    <xsl:variable name="props" as="item()*">
      <xsl:apply-templates select="$lvl/w:pPr" mode="#current"/>
    <xsl:apply-templates select="for $style in reverse($based-on-chain/*[position() le $numId-chain-item-pos])
                                   return $style/w:pPr/*[not(self::w:numPr)]" mode="#current"/>
    </xsl:variable>
    <xsl:sequence select="$props//docx2hub:attribute[starts-with(@name, 'css:')]"/>
    <xsl:next-match>
      <xsl:with-param name="ilvl" as="xs:integer" select="$ilvl" tunnel="yes"/>
    </xsl:next-match>
  </xsl:template>
  
  <xsl:template match="w:style/w:pPr/w:numPr[not(w:ilvl)]" mode="docx2hub:add-props" priority="1">
    <xsl:param name="ilvl" as="xs:integer?" tunnel="yes"/>
    <xsl:copy>
      <xsl:apply-templates select="*" mode="#current"/>
      <xsl:if test="$ilvl">
        <w:ilvl w:val="{$ilvl}"/>  
      </xsl:if>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:lvl/w:rPr | w:lvl/w:pPr" mode="docx2hub:add-props" priority="3">
    <xsl:copy>
      <xsl:apply-templates select="*" mode="#current" />
      <!-- in order for subsequent (numbering.xsl) symbol mappings, the original rFonts must also be preserved -->
      <xsl:sequence select="*"/>
    </xsl:copy>
  </xsl:template>
  
  <!-- § 17.3.1.29: this is only for the paragraph mark's formatting: -->
  <xsl:template match="w:pPr/w:rPr" mode="docx2hub:add-props" priority="2.5" />

  <xsl:template match="w:tblPr | w:tblPrEx" mode="docx2hub:add-props" priority="2">
    <xsl:copy>
      <xsl:apply-templates select="*" mode="#current" />
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:tblPrEx" mode="docx2hub:add-props">
    <xsl:apply-templates select="* except w:tblW" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:trPr" mode="docx2hub:add-props">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <xsl:template match="w:rubyPr" mode="docx2hub:add-props">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <xsl:template match="w:tcPr" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="*" mode="#current" />
    <!-- for cellspan etc. processing as defined in tables.xsl: -->
    <xsl:sequence select="." />
  </xsl:template>

  <xsl:template match="w:u" mode="docx2hub:add-props" priority="2">
    <xsl:next-match/>
    <xsl:apply-templates select="@w:color" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:ind" mode="docx2hub:add-props" priority="2">
    <!-- Precedence as per § 17.3.1.12.
         Ignore the ..Chars variants and start/end for the time being. -->
    <xsl:apply-templates select="@w:left | @w:right | @w:firstLine, @w:hanging" mode="#current" />
  </xsl:template>

  <xsl:template match="w:pBdr | w:tcBorders | w:tblBorders | w:tblCellMar | w:tcMar" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="w:left | w:right | w:top | w:bottom | w:insideH | w:insideV" mode="#current" />
  </xsl:template>

  <xsl:template match="w:spacing" mode="docx2hub:add-props" priority="2">
    <!-- Precedence as per § 17.3.1.33.
         Ignore autospacing.
         Handles both w:pPr/w:spacing and w:rPr/w:spacing in this template rule. -->
    <xsl:apply-templates select="@w:val | @w:after | @w:before | @w:line, @w:afterLines | @w:beforeLines" mode="#current" />
  </xsl:template>

  <xsl:template match="w:tcW" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="@w:w" mode="#current" />
  </xsl:template>

  <!-- ISO/IEC 29500-1 (2008-11-15):
       „If this run has any background shading specified using the shd element (§17.3.2.32), 
       then the background shading shall be superseded by the highlighting color“ -->
  <xsl:template match="w:shd[../w:highlight]" mode="docx2hub:add-props" priority="2"/>

  <xsl:template match="w:sectPr[parent::w:pPr]" mode="docx2hub:add-props" priority="+2">
    <xsl:apply-templates select="w:pgSz | w:footnotePr" mode="#current"/>
  </xsl:template>
  
  <!-- As a collateral, move column/page breaking instructions to the previous para and 
    mark the para that they’ve been in as removable, because these paras won’t be displayed by 
    Word. -->
  <xsl:function name="docx2hub:is-removable-para" as="xs:boolean">
    <xsl:param name="para" as="element(w:p)"/>
    <xsl:sequence select="exists($para[not(.//w:r)]/w:pPr/w:sectPr)
                          or
                          exists($para[w:r/w:br[@w:type = 'page']]
                                      [every $c in w:r/* satisfies $c/self::w:br[@w:type = 'page']])"/>
  </xsl:function>

  <xsl:template match="w:p" mode="docx2hub:add-props">
    <!-- Pseudo paras that are meant to hold only sectPr or page breaks will be removed in a later pass -->
    <xsl:copy>
      <xsl:variable name="is-removable" select="docx2hub:is-removable-para(.)" as="xs:boolean"/>
      <xsl:if test="$is-removable">
        <xsl:attribute name="docx2hub:removable" select="$is-removable"/>
      </xsl:if>
      <xsl:apply-templates select="@*, *" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="  w:style/*
                       | w:rPr[not(ancestor::m:oMath)]/* 
                       | w:pBdr/* 
                       | w:pPr/* 
                       | w:tblPr/*[not(self::w:jc)]
                       | w:tcBorders/* 
                       | w:tblBorders/*
                       | w:tcPr/*
                       | w:trPr/*[not(local-name() = ('tblHeader','jc'))]
                       | w:rubyPr/*
                       | w:tblPrEx/*
                       | w:tblCellMar/*
                       | w:tcMar/*
                       | w:pgSz/@*
                       | w:ind/@* 
                       | w:tab/@*[local-name() ne 'srcpath']
                       | w:tcW/@* 
                       | w:u/@w:color
                       | w:spacing/@* 
                       | v:shape/@* 
                       | wp:extent/@*
                       | m:oMathParaPr/*
                       " 
    mode="docx2hub:add-props">
    <xsl:variable name="prop" select="key('docx2hub:prop', docx2hub:propkey(.), $docx2hub:propmap)" />
    <xsl:variable name="raw-output" as="element(*)*">
      <xsl:apply-templates select="$prop" mode="#current">
        <xsl:with-param name="val" select="." tunnel="yes" as="item()"/>
      </xsl:apply-templates>
      <xsl:if test="empty($prop)">
        <!-- Fallback (no mapping in propmap): -->
        <docx2hub:attribute name="docx2hub:generated-{local-name()}"><xsl:value-of select="docx2hub:serialize-prop(.)" /></docx2hub:attribute>
      </xsl:if>
    </xsl:variable>
    <xsl:sequence select="$raw-output" />
  </xsl:template>
  
  <xsl:template match="w:tcW[@w:type = 'pct']/@w:w" mode="docx2hub:add-props">
    <xsl:variable name="pct-to-dxa" as="element()">
      <xsl:variable name="tbl-w" select="../../../../../w:tblPr/w:tblW/@w:w" as="xs:decimal"/>
      <xsl:variable name="tc-pct" select="if(matches(., '%'))
        then xs:double(replace(., '%', ''))
        else xs:double(.) div 5000" as="xs:double"/>
      <xsl:variable name="tc-dxa" select="$tbl-w * $tc-pct" as="xs:double"/>
      <w:tcW w:type="dxa" w:w="{round-half-to-even($tc-dxa)}"/>
    </xsl:variable>
    <xsl:apply-templates select="$pct-to-dxa" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:tblPr/w:jc" mode="docx2hub:add-props">
    <xsl:variable name="tbl-ind" as="node()*">
      <xsl:if test="exists(parent::w:tblPr/w:tblInd)">
        <xsl:apply-templates select="parent::w:tblPr/w:tblInd" mode="#current"/>
      </xsl:if>
    </xsl:variable>
    <docx2hub:attribute name="css:margin-left">
      <xsl:value-of select="if (exists($tbl-ind) and not($tbl-ind//text()='')) then $tbl-ind//text() else if (@w:val=('left','start')) then '0pt' else 'auto'"/>
    </docx2hub:attribute>
    <docx2hub:attribute name="css:margin-right">
      <xsl:value-of select="if (@w:val=('right','end')) then '0pt' else 'auto'"/>
    </docx2hub:attribute>
  </xsl:template>

  <xsl:template match="@w:rsid
                       | @w:rsidDel
                       | @w:rsidR
                       | @w:rsidRPr
                       | @w:rsidRDefault
                       | @w:rsidP
                       | @w:rsidTr
                       | @w14:paraId
                       | @w14:textId"
    mode="docx2hub:add-props" />

  <xsl:function name="docx2hub:propkey" as="xs:string">
    <xsl:param name="prop" as="node()" /> <!-- w:sz, ... -->
    <xsl:choose>
      <xsl:when test="$prop/self::attribute()">
        <xsl:sequence select="string-join((name($prop/..), name($prop)), '/@')" />
      </xsl:when>
      <xsl:when test="$prop/(parent::w:pBdr, parent::w:tcBorders, parent::w:tblBorders, parent::w:tblCellMar, parent::w:tcMar)">
        <xsl:sequence select="string-join((name($prop/..), name($prop)), '/')" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="name($prop)" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="docx2hub:serialize-prop" as="xs:string">
    <xsl:param name="prop" as="item()" /> <!-- w:sz, @fillcolor, ... -->
    <xsl:sequence select="string-join(
                            for $a in 
                              if ($prop instance of element()) 
                              then $prop/@* 
                              else $prop (: attribute() :)
                            return concat(name($a), ':', $a),
                            '; '
                          )" />
  </xsl:function>

  <xsl:template match="prop" mode="docx2hub:add-props" as="node()*">
    <xsl:variable name="atts" as="element(*)*">
      <!-- in the following line, val is a potential child of prop (do not cofuse with $val)! -->
      <xsl:apply-templates select="@type, val, @target-value[not(../(@type, val))]" mode="#current" />
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="empty($atts) and @default">
        <docx2hub:attribute name="{@target-name}"><xsl:value-of select="@default" /></docx2hub:attribute>
      </xsl:when>
      <xsl:when test="empty($atts)" />
      <xsl:otherwise>
        <xsl:sequence select="$atts" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="prop[not(@type | val)]/@target-value" mode="docx2hub:add-props">
    <xsl:call-template name="docx2hub:XML-Hubformat-atts"/>
  </xsl:template>

  <xsl:template match="val" mode="docx2hub:add-props" as="element(*)?">
    <xsl:apply-templates select="@eq, @match" mode="#current" />
  </xsl:template>

  <xsl:template match="prop/@type" mode="docx2hub:add-props" as="node()*">
    <xsl:param name="val" as="item()" tunnel="yes" /><!-- element or attribute -->
    <xsl:choose>

      <xsl:when test=". eq 'docx-line'">
        <xsl:choose>
          <xsl:when test="$val/parent::*[@w:lineRule and not(@w:lineRule='auto')]">
            <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="docx2hub:pt-length(($val, $val/@w:val, $val/@w:w)[normalize-space()][1])" /></docx2hub:attribute>
          </xsl:when>
          <xsl:otherwise>
            <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="number($val) div 240" /></docx2hub:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test=". eq 'percentage'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="if (xs:integer($val) eq -1) then 1 else xs:double($val) * 0.01" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'lang'">
        <xsl:variable name="stringval" as="xs:string?" 
          select="if ($val/self::w:lang) then $val/@w:val[1] else $val"/>
        <!-- stripping the country by default. If someone needs it, we need to introduce an option -->
        <xsl:variable name="repl" as="xs:string" 
          select="if (matches($stringval, 'German') or matches($stringval, '\Wde\W'))
                  then 'de'
                  else 
                    if (matches($stringval, 'English'))
                    then 'en'
                    else replace($stringval, '^(\p{Ll}+).*$', '$1')" />
        <xsl:if test="normalize-space($repl)">
          <docx2hub:attribute name="{../@target-name}">
            <xsl:value-of select="$repl"/>  
          </docx2hub:attribute>
        <!-- stripping the country by default. If someone needs it, we need to introduce an option -->
        </xsl:if>
      <xsl:if test="$val/self::w:lang[not(@w:val)]/@w:bidi">
          <docx2hub:attribute name="docx2hub:rtl-lang">
            <xsl:value-of select="replace($val/@w:bidi, '^(\p{Ll}+).*$', '$1')"/>  
          </docx2hub:attribute>
        </xsl:if>
      </xsl:when>

      <xsl:when test=". eq 'docx-boolean-prop'">
        <xsl:choose>
          <xsl:when test="$val/@w:val = ('0','false') and exists(../@default)">
            <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="../@default" /></docx2hub:attribute>
          </xsl:when>
          <xsl:otherwise>
            <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="../@active" /></docx2hub:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test=". eq 'docx-bdr'">
        <xsl:variable name="borders" as="element(w:pBdr)">
          <w:pBdr>
            <xsl:for-each select="('left', 'right', 'bottom', 'top')">
              <xsl:element name="w:{.}">
                <xsl:sequence select="$val/@*"/>
              </xsl:element>
            </xsl:for-each>
          </w:pBdr>
        </xsl:variable>
        <xsl:apply-templates select="$borders/*" mode="#current"/>
      </xsl:when>
      
      <xsl:when test=". eq 'docx-border'">
        <!-- According to § 17.3.1.5 and other sections, the top/bottom borders don't apply
             if a set of paras has identical border settings. The between setting should be used instead.
             TODO -->
        <xsl:variable name="orientation" select="replace(../@name, '^.+:', '')" as="xs:string"/>
        <docx2hub:attribute name="css:border-{$orientation}-style">
          <xsl:value-of select="docx2hub:border-style($val/@w:val)"/>
        </docx2hub:attribute>
        <xsl:if test="$val/@w:val and not($val/@w:val = ('nil','none'))">
          <docx2hub:attribute name="css:border-{$orientation}-width">
            <xsl:value-of select="docx2hub:pt-border-size($val/@w:sz)"/>
          </docx2hub:attribute>
          <xsl:if test="$val/@w:color ne 'auto'">
            <docx2hub:attribute name="css:border-{$orientation}-color">
              <xsl:value-of select="docx2hub:color($val/@w:color)"/>
            </docx2hub:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      
      <xsl:when test=". eq 'docx-padding'">
        <xsl:variable name="orientation" select="replace(../@name, '^.+:', '')" as="xs:string" />
        <docx2hub:attribute name="css:padding-{$orientation}">
          <!-- LibreOffice produced a padding of -2 dxa, so check for negativity -->
          <xsl:value-of select="if (starts-with($val/@w:w, '-')) then '0' else docx2hub:pt-length($val/@w:w)" />
        </docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-charstyle'">
        <xsl:variable name="linked" as=" xs:string?" select="key('docx2hub:style', $val/@w:val, root($val))/w:link/@w:val"/>
        <xsl:call-template name="docx2hub:style-name">
          <xsl:with-param name="val" select="$val"/>
          <xsl:with-param name="linked" select="$linked"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test=". eq 'docx-color'">
        <xsl:variable name="colorval" as="xs:string?" select="docx2hub:color( ($val/@w:val, $val)[1] )" />
        <xsl:if test="$colorval">
          <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$colorval" /></docx2hub:attribute>
        </xsl:if>
      </xsl:when>

      <xsl:when test=". eq 'docx-font-family'">
        <xsl:if test="$val/@w:ascii or $val/@w:hAnsi">
          <!-- We should implement the complex font selection rules as defined in 
            https://onedrive.live.com/view.aspx/Public%20Documents/2009/DR-09-0040.docx?cid=c8ba0861dc5e4adc&sc=documents -->
          <xsl:variable name="font-name" select="($val/@w:ascii, $val/@w:hAnsi)[1]" as="xs:string"/>
          <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$font-name"/></docx2hub:attribute>
          <xsl:variable name="charset" as="xs:string*" 
            select="distinct-values(key('docx2hub:font-by-name', $font-name, root($val))/w:charset/@w:val)"/>
          <!-- xs:string* instead of xs:string? because there are docx files with multiple w:font entries for a given @w:name -->
          <xsl:if test="exists($charset) 
                        and not($charset = ('0', '00', '02', '80', '81', '86', 'EE')) (: 80: Arial Unicode MS, MS Mincho, … :)
                        and not($val/ancestor::w:style)">
            <!-- I saw 'C8' for SMinionPlus. Don’t know whether this may be treated as Unicode. Probably not,
            since it may contain variants of the Springer logo at various positions. We should supply a mapping 
            for this. Also need to establish a mechanism for the situation when mosts characters of a font map to Unicode except 
            for a few exceptions. -->
            <!-- mk: I recently saw a document with value '81' for MS Arial Unicode. Probably this value is permitted too -->
            <docx2hub:attribute name="docx2hub:map-from"><xsl:value-of select="$font-name"/></docx2hub:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:when>

      <xsl:when test=". eq 'docx-font-size'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="concat(number($val/@w:val) * 0.5, 'pt')" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-font-stretch'">
        <xsl:variable name="result" as="xs:string">
          <xsl:choose>
            <xsl:when test="$val/@w:val &lt; 40">
              <xsl:sequence select="'ultra-condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 60">
              <xsl:sequence select="'extra-condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 80">
              <xsl:sequence select="'condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 96">
              <xsl:sequence select="'semi-condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 104">
              <xsl:sequence select="'normal'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 120">
              <xsl:sequence select="'semi-expanded'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 140">
              <xsl:sequence select="'expanded'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 160">
              <xsl:sequence select="'extra-expanded'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:sequence select="'ultra-expanded'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$result" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-hierarchy-level'">
        <xsl:if test="not($val/@w:val eq '9')"><!-- § 17.3.1.20 -->
          <docx2hub:attribute name="remap"><xsl:value-of select="concat('h', xs:integer($val/@w:val) + 1)" /></docx2hub:attribute>
        </xsl:if>
      </xsl:when>

      <xsl:when test=". eq 'docx-length-attr'">
        <xsl:if test="not($val/../@w:w='0' and $val/../@w:type='auto')">
          <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="docx2hub:pt-length(($val, $val/@w:val, $val/@w:w)[normalize-space()][1])" /></docx2hub:attribute>
        </xsl:if>
      </xsl:when>
      
      <xsl:when test=". eq 'docx-image-size-attr'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="concat(number($val) div 12700,'pt')" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-length-attr-negated'">
        <xsl:variable name="string-val" select="string(($val/@w:val, $val)[1])" as="xs:string"/>
        <docx2hub:attribute name="{../@target-name}">
          <xsl:value-of select="if (matches($string-val, '^-'))
                                then docx2hub:pt-length(replace($string-val, '^-', ''))
                                else docx2hub:pt-length(concat('-', $string-val))" />
        </docx2hub:attribute>
      </xsl:when>
      
      <xsl:when test=". eq 'docx-position-attr-negated'">
        <xsl:variable name="string-val" select="string(number(($val/@w:val, $val)[1]) * 10)" as="xs:string"/>
        <docx2hub:attribute name="{../@target-name}">
          <xsl:value-of select="if (matches($string-val, '^-'))
                                  then docx2hub:pt-length(replace($string-val, '^-', ''))
                                  else docx2hub:pt-length(concat('-', $string-val))" />
        </docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-parastyle'">
        <xsl:call-template name="docx2hub:style-name">
          <xsl:with-param name="val" select="$val"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test=". eq 'docx-shd'">
        <xsl:choose>
          <xsl:when test="$val/@w:val = ('clear','nil')">
            <xsl:if test="$val/@w:fill ne 'auto'">
              <docx2hub:attribute name="css:background-color"><xsl:value-of select="concat('#', $val/@w:fill)" /></docx2hub:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$val/@w:val eq 'solid'">
            <docx2hub:attribute name="css:background-color"><xsl:value-of select="concat('#', $val/@w:color)" /></docx2hub:attribute>
          </xsl:when>
          <xsl:when test="matches($val/@w:val, '^pct')">
            <xsl:choose>
              <xsl:when test="exists($val/@w:fill) and exists($val/@w:color)">
                <xsl:if test="not(matches($val/@w:color,'auto'))">
                  <docx2hub:attribute name="css:color"><xsl:value-of select="concat('#', $val/@w:color)"/></docx2hub:attribute>
                </xsl:if>
                <docx2hub:color-percentage target="css:background-color" use="css:color" fill="{if ($val/@w:fill='auto') then '#FFFFFF' else concat('#',$val/@w:fill)}"><xsl:value-of select="replace($val/@w:val, '^pct', '')" /></docx2hub:color-percentage>
              </xsl:when>
              <xsl:otherwise>
                <xsl:message>map-props.xsl: w:shd/@w:val='pct*' only implemented for existing @w:fill and @w:color
                <xsl:sequence select="$val" />
              </xsl:message>
            </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:message>map-props.xsl: w:shd/@w:val other than 'clear', 'nil', 'pct*', and 'solid' not implemented.
            <xsl:sequence select="$val" />
            </xsl:message>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test=". eq 'docx-table-row-height'">
        <xsl:variable name="pt-length" select="docx2hub:pt-length($val/@w:val)" as="xs:string"/>
        <xsl:choose>
          <xsl:when test="$val/@w:hRule = 'auto'"/>
          <xsl:when test="$val/@w:hRule = 'atLeast' or (not($val/@w:hRule))">
            <docx2hub:attribute name="css:min-height"><xsl:value-of select="$pt-length" /></docx2hub:attribute>    
          </xsl:when>
          <xsl:when test="$val/@w:hRule = 'exact' 
                          (:or 
                          (
                            not($val/@w:hRule)
                            and
                            $val/ancestor::w:tbl[1]/w:tblPr/w:tblLayout/@w:type = 'fixed'
                          ) :)">
            <!-- not sure about the last condition. § 17.4.81 says:
              “If [@w:hRule] is omitted, then its value shall be assumed to be auto.”
              But there were table rows with a trHeight that lacked @w:hRule, and their height was fixed.
              They were in a fixed-layout table, so we assume that row heights should be respected (at least) in
              fixed tables even if their @w:hRule is missing. -->
            <docx2hub:attribute name="css:height"><xsl:value-of select="$pt-length" /></docx2hub:attribute>    
          </xsl:when>
        </xsl:choose>
        
      </xsl:when>
      
      <xsl:when test=". eq 'docx-underline'">
        <!-- §§§ TODO -->
        <docx2hub:attribute name="css:text-decoration-line"><xsl:value-of select="'underline'" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'tablist'">
        <tabs>
          <xsl:apply-templates select="$val/*" mode="#current"/>
        </tabs>
      </xsl:when>

      <xsl:when test=". eq 'docx-text-direction'">
        <!-- I find 17.18.93 ST_TextDirection remarkably unclear about this. What also bothers me is that
             the value 'btLr' doesn’t appear in the table in that section. In Annex N.1 on p. 5563, they mention
             that btLr et al. have been dropped in Wordprocessing ML. -->
        <xsl:choose>
          <xsl:when test="$val/@w:val = ('tbLr', 'btLr')">
            <!-- preliminary value – only works in IE, while the CSS3 writing mode prop values don’t work -->
            <docx2hub:attribute name="css:writing-mode">bt-lr</docx2hub:attribute>
          </xsl:when>
          <xsl:when test="matches($val/@w:val, 'tb', 'i')">
            <!-- looks funny -->
            <docx2hub:attribute name="css:transform">rotate(90deg)</docx2hub:attribute>
          </xsl:when>
          <xsl:when test="matches($val/@w:val, 'bt', 'i')">
            <!-- looks funny -->
            <docx2hub:attribute name="css:transform">rotate(-90deg)</docx2hub:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:message>Unsupported text direction property <xsl:sequence select="$val"/></xsl:message>
          </xsl:otherwise>
        </xsl:choose>
        <!-- no effect: -->
        <!--<docx2hub:attribute name="css:width">fit-content</docx2hub:attribute>-->
      </xsl:when>
      
      <xsl:when test=". eq 'linear'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="if ($val/self::attribute())
                                                                           then $val
                                                                           else $val/@w:val" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'passthru'">
        <xsl:copy-of select="$val" copy-namespaces="no"/>
      </xsl:when>

      <xsl:when test=". eq 'docx-position'">
        <xsl:choose>
          <xsl:when test="$val/@w:val eq 'baseline'" />
          <xsl:otherwise>
            <docx2hub:wrap element="{$val/@w:val}" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test=". eq 'style-link'">
        <docx2hub:style-link type="{../@name}" target="{$val}"/>
      </xsl:when>

      <xsl:when test=". eq 'style-name'">
        <docx2hub:attribute name="{../@target-name}">
          <xsl:value-of select="docx2hub:css-compatible-name($val/@w:val)"/>
        </docx2hub:attribute>
      </xsl:when>

      <xsl:otherwise>
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$val" /></docx2hub:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- ISO/IEC 29500-1, 17.3.2.40: 
    This color therefore can be automatically be modified by a consumer as appropriate, 
    for example, in order to ensure that the underline can be distinguished against the 
    page's background color. --> 
  <xsl:variable name="docx2hub:auto-color" select="'#000000'" as="xs:string"/>

  <xsl:function name="docx2hub:color" as="xs:string?" >
    <xsl:param name="val" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$val = 'none'" />
      <xsl:when test="$val = 'auto'" >
        <!-- not tested yet whether this interferes with <w:shd w:fill="…"/> -->
        <xsl:sequence select="$docx2hub:auto-color" />
      </xsl:when>
      <xsl:when test="matches($val, '^#[0-9a-f]{6}$')">
        <!-- e.g., v:shape/@fillcolor -->
        <xsl:sequence select="upper-case($val)" />
      </xsl:when>
      <xsl:when test="matches($val, '[0-9A-F]{6}')">
        <xsl:sequence select="concat('#', $val)" />
      </xsl:when>
      <xsl:when test="$val eq 'cyan'">
        <xsl:sequence select="'#00FFFF'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkBlue'">
        <xsl:sequence select="'#00008B'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkCyan'">
        <xsl:sequence select="'#008B8B'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkGray'">
        <xsl:sequence select="'#A9A9A9'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkGreen'">
        <xsl:sequence select="'#006400'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkMagenta'">
        <xsl:sequence select="'#800080'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkRed'">
        <xsl:sequence select="'#8B0000'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkYellow'">
        <xsl:sequence select="'#808000'" />
      </xsl:when>
      <xsl:when test="$val eq 'green'">
        <xsl:sequence select="'#00FF00'" />
      </xsl:when>
      <xsl:when test="$val eq 'lightGray'">
        <xsl:sequence select="'#D3D3D3'" />
      </xsl:when>
      <xsl:when test="$val eq 'magenta'">
        <xsl:sequence select="'#FF00FF'" />
      </xsl:when>
      <xsl:when test="$val eq 'yellow'">
        <xsl:sequence select="'#FFFF00'" />
      </xsl:when>
      <xsl:when test="$val eq 'blue'">
        <xsl:sequence select="'#0000FF'" />
      </xsl:when>
      <xsl:when test="$val eq 'red'">
        <xsl:sequence select="'#FF0000'" />
      </xsl:when>
      <xsl:when test="$val eq 'black'">
        <xsl:sequence select="'#000000'" />
      </xsl:when>
      <xsl:when test="$val eq 'white'">
        <xsl:sequence select="'#FFFFFF'" />
      </xsl:when>
      <xsl:otherwise><!-- shouldn't happen -->
        <xsl:sequence select="''" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template name="docx2hub:style-name" as="element(docx2hub:attribute)">
    <xsl:param name="val" as="element(*)"/><!-- w:pStyle, w:cStyle -->
    <xsl:param name="linked" as="xs:string?"/>
    <!-- we choose not to use the linked (paragraph) style here because we’d have to
      carefully select only the css:rule’s inline properties when recreating docx run properties
      from hub. -->
    <xsl:variable name="looked-up" as="xs:string" select="$val/@w:val" />
    <!--<xsl:variable name="looked-up" as="xs:string" select="if ($linked) then $linked else $val/@w:val" />-->
    <docx2hub:attribute name="role">
      <xsl:value-of select="if ($hub-version eq '1.0')
                              then $looked-up
                              else docx2hub:css-compatible-name($looked-up)" />
    </docx2hub:attribute>
  </xsl:template>

  <xsl:function name="docx2hub:pt-length" as="xs:string" >
    <xsl:param name="val" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="not($val)">
        <xsl:message>empty argument for docx2hub:pt-length, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:when test="not($val castable as xs:integer)">
        <xsl:message>argument '<xsl:value-of select="$val"/>' for docx2hub:pt-length not castable as xs:integer, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="if (matches($val, '%$'))
          then $val
          else concat(xs:string(xs:integer($val) * 0.05), 'pt')" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="docx2hub:pt-border-size" as="xs:string" >
    <xsl:param name="val" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="not($val)">
        <xsl:message>empty argument for docx2hub:pt-border-size, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:when test="not($val castable as xs:integer)">
        <xsl:message>argument '<xsl:value-of select="$val"/>' for docx2hub:border-size not castable as xs:integer, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="if (matches($val, '%$'))
          then $val
          else concat(xs:string(min((12, max((0.25, xs:integer($val) * 0.125))))), 'pt')" />
        <!-- 17.3.4 minimum:0.25pt, maximum: 12pt -->
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:function name="docx2hub:border-style" as="xs:string" >
    <xsl:param name="val" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="$val eq 'single'">
        <xsl:sequence select="'solid'" />
      </xsl:when>
      <xsl:when test="matches($val, '(thinThick.+Gap|dashDotStroked)')">
        <xsl:sequence select="'dashed'" />
      </xsl:when>
      <xsl:when test="$val eq 'threeDEmboss'">
        <xsl:sequence select="'groove'" />
      </xsl:when>
      <xsl:when test="$val eq 'nil' or not($val)">
        <xsl:sequence select="'none'" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$val" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="val/@match" mode="docx2hub:add-props" as="element(*)?">
    <xsl:param name="val" as="item()" tunnel="yes" />
    <xsl:if test="matches($val/@*:val, .)">
      <xsl:call-template name="docx2hub:XML-Hubformat-atts" />
    </xsl:if>
  </xsl:template>

  <xsl:template match="val/@eq" mode="docx2hub:add-props" as="element(*)?">
    <xsl:param name="val" as="item()" tunnel="yes" />
    <xsl:if test="string($val) = string(.) or (string(.) = 'true' and empty($val/@*))">
      <xsl:call-template name="docx2hub:XML-Hubformat-atts" />
    </xsl:if>
  </xsl:template>

  <xsl:template name="docx2hub:XML-Hubformat-atts" as="element(*)?">
    <xsl:variable name="target-val" select="(../@target-value, ../../@target-value)[last()]" as="xs:string?" />
    <xsl:if test="exists($target-val)">
      <docx2hub:attribute name="{(../@target-name, ../../@target-name)[last()]}"><xsl:value-of select="$target-val" /></docx2hub:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template match="Color" mode="docx2hub:add-props" as="xs:string">
    <xsl:param name="multiplier" as="xs:double" select="1.0" />
    <xsl:choose>
      <xsl:when test="@Space eq 'CMYK'">
        <xsl:sequence select="concat(
                                'device-cmyk(', 
                                string-join(
                                  for $v in tokenize(@ColorValue, '\s') return xs:string(xs:integer(xs:double($v) * 10000 * $multiplier) * 0.000001)
                                  , ','
                                ),
                                ')'
                              )" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:message>Unknown colorspace <xsl:value-of select="@Space"/>
        </xsl:message>
        <xsl:sequence select="@ColorValue" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:tabs/w:tab" mode="docx2hub:add-props" as="element(dbk:tab)" priority="2">
    <tab>
      <xsl:apply-templates select="@*" mode="#current" />
    </tab>
  </xsl:template>

  <xsl:template match="w:tabs/w:tab[@w:val eq 'clear']" mode="docx2hub:add-props" priority="3"/>


  <xsl:key name="docx2hub:style" 
    match="CellStyle | CharacterStyle | ObjectStyle | ParagraphStyle | TableStyle" 
    use="@Self" />

  <xsl:function name="docx2hub:css-compatible-name" as="xs:string">
    <xsl:param name="input" as="xs:string"/>
    <xsl:sequence select="replace(  
                            replace(
                              replace(
                                normalize-unicode($input, 'NFKD'), 
                                '\p{Mn}', 
                                ''
                              ), 
                              '[^-_a-z0-9]', 
                              '_', 
                              'i'
                            ),
                            '^(\I)',
                            '_$1'
                          )"/>
  </xsl:function>

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- mode: docx2hub:props2atts -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template match="* | @*" mode="docx2hub:props2atts">
    <xsl:variable name="content" as="node()*">
      <xsl:apply-templates select="docx2hub:style-link" mode="#current" />
      <xsl:apply-templates select="m:oMathPara/docx2hub:attribute[not(@name = ../docx2hub:remove-attribute/@name)]" mode="#current" />
      <xsl:apply-templates select="docx2hub:attribute[not(@name = following-sibling::docx2hub:remove-attribute/@name)]" mode="#current" />
      <xsl:apply-templates select="docx2hub:color-percentage[not(@name = following-sibling::docx2hub:remove-attribute/@name)]" mode="#current" />
      <xsl:variable name="remaining-tabs" as="element(dbk:tab)*">
        <xsl:for-each-group select="dbk:tabs/dbk:tab" group-by="docx2hub:attribute[@name eq 'horizontal-position']">
          <xsl:if test="not(current-group()[last()]/docx2hub:attribute[@name eq 'clear'])">
            <xsl:apply-templates select="current-group()[last()]" mode="#current"/>
          </xsl:if>
        </xsl:for-each-group>
      </xsl:variable>
      <xsl:if test="exists($remaining-tabs)">
        <tabs>
          <xsl:sequence select="$remaining-tabs" />
        </tabs>
      </xsl:if>
      <xsl:apply-templates select="node() except (docx2hub:attribute | docx2hub:color-percentage | docx2hub:wrap | docx2hub:style-link | dbk:tabs)" mode="#current" />
    </xsl:variable>
    <xsl:choose>
      <!-- do not wrap whitespace only subscript or superscript -->
      <xsl:when test="w:t and docx2hub:wrap/@element = ('superscript', 'subscript') 
                      and not(exists(docx2hub:wrap/@element[ . ne 'superscript' and . ne 'subscript']))
                      and (
                        every $el in $content[self::*] 
                        satisfies $el[self::w:t[@xml:space eq 'preserve'][matches(., '^\p{Zs}*$')]]
                      )">
        <xsl:copy>
          <xsl:sequence select="docx2hub:wrap((@srcpath, $content), (docx2hub:wrap[not(@element = ('superscript', 'subscript'))]))" />
        </xsl:copy>
      </xsl:when>
      <!-- do not wrap field function elements in subscript or superscript.
      Exception: instrText will be wrapped (see next xsl:when) when it isn’t the first instrText after a 'begin' fldChar.
      There may be sub/superscripts in index terms. However, sometimes even the field function name (XE for index terms)
      is lowered or raised. This leads to errors when processing it. -->
      <xsl:when test="docx2hub:wrap/@element = ('superscript', 'subscript') 
                      and
                      (exists(w:fldChar | w:instrText[current()/preceding-sibling::*[1]/self::w:r/w:fldChar[@w:fldCharType = 'begin']] ))
                      and
                      (every $i in * satisfies $i[self::w:fldChar (:or self::w:instrText:) or self::docx2hub:*])">
        <xsl:copy>
          <xsl:sequence select="docx2hub:wrap((@srcpath, $content), (docx2hub:wrap[not(@element = ('superscript', 'subscript'))]))" />
        </xsl:copy>
      </xsl:when>
      <!-- Deal with sub/sup in index terms: -->
      <xsl:when test="docx2hub:wrap/@element = ('superscript', 'subscript') 
                      and
                      (exists(w:instrText))
                      and
                      (every $i in * satisfies exists($i/(self::w:instrText | self::docx2hub:*)))">
        <xsl:copy>
          <xsl:apply-templates select="docx2hub:attribute" mode="#current"/>
          <xsl:for-each select="w:instrText">
            <xsl:copy>
              <xsl:copy-of select="@*"><!-- @xml:space --></xsl:copy-of>
              <xsl:sequence select="docx2hub:wrap((../@srcpath, node()), (../docx2hub:wrap))" />
            </xsl:copy>
          </xsl:for-each>
        </xsl:copy>
      </xsl:when>
      <xsl:when test="exists(docx2hub:wrap) and exists(self::css:rule | self::dbk:style)">
        <xsl:copy>
          <xsl:attribute name="remap" select="docx2hub:wrap/@element" />
          <xsl:sequence select="@*, (@srcpath, $content)" />
        </xsl:copy>
      </xsl:when>
      <xsl:when test="exists(docx2hub:wrap) and not(self::css:rule or self::dbk:style)">
        <xsl:sequence select="docx2hub:wrap((@srcpath, $content), (docx2hub:wrap))" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:copy>
          <xsl:sequence select="@*, $content" />
        </xsl:copy>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="docx2hub:wrap" as="node()*">
    <xsl:param name="content" as="item()*" /><!-- attribute(srcpath) or node() -->
    <xsl:param name="wrappers" as="element(docx2hub:wrap)*" />
    <xsl:choose>
      <xsl:when test="exists($wrappers)">
        <xsl:element name="{$wrappers[1]/@element}">
          <xsl:sequence select="docx2hub:wrap($content, $wrappers[position() gt 1])" />
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="$content" mode="docx2hub:props2atts"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="docx2hub:attribute" mode="docx2hub:props2atts">
    <xsl:attribute name="{@name}" select="." />
  </xsl:template>

  <xsl:template match="docx2hub:attribute[@name = ('fill-tint')]" mode="docx2hub:props2atts"/>

  <xsl:template match="docx2hub:attribute[@name = 'css:letter-spacing'] 
                                         [not(parent::css:rule)]
                                         [not(../docx2hub:attribute[@name eq 'role'])]
                                         [. = '0pt']" mode="docx2hub:props2atts"/>

  <xsl:template match="docx2hub:attribute[@name='role']
                                         [parent::w:p[w:pgSz]
                                                     [every $i 
                                                      in child::* 
                                                      satisfies $i[not(self::w:r) and 
                                                                   not(self::m:oMathPara) and 
                                                                   not(self::m:oMath) and 
                                                                   not(self::w:fldSimple) and 
                                                                   not(self::w:hyperlink)]]]" mode="docx2hub:props2atts"/>
  
  <xsl:template match="docx2hub:attribute[@name = 'css:text-decoration-line']" mode="docx2hub:props2atts">
    <xsl:variable name="all-atts" select="preceding-sibling::docx2hub:attribute[@name = current()/@name], ."
      as="element(docx2hub:attribute)+"/>
    <xsl:variable name="tokenized" select="for $a in $all-atts return tokenize($a, '\s+')" as="xs:string+"/>
    <xsl:variable name="line-through" select="$tokenized[starts-with(., 'line-through')][last()]"/>
    <xsl:variable name="underline" select="$tokenized[starts-with(., 'underline')][last()]"/>
    <xsl:choose>
      <xsl:when test="every $t in ($line-through, $underline) satisfies (ends-with($t, 'none'))">
        <xsl:attribute name="{@name}" select="'none'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="{@name}" select="($line-through, $underline)[not(ends-with(., 'none'))]"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="docx2hub:color-percentage" mode="docx2hub:props2atts">
    <xsl:variable name="color" select="(../docx2hub:attribute[@name eq 'css:color'][last()], '#000000')[1]" as="xs:string" />
    <xsl:attribute name="{@target}" select="tr:tint-hex-color-filled($color, number(.) * 0.01, @fill)" />
  </xsl:template>

  <!-- Let only the last setting prevail: -->
  <xsl:template match="w:numPr[following-sibling::w:numPr]" mode="docx2hub:props2atts" />
  <xsl:template match="w:tblPr[following-sibling::w:tblPr]" mode="docx2hub:props2atts" />
  <xsl:template match="w:tcPr[following-sibling::w:tcPr]" mode="docx2hub:props2atts" />

  <xsl:template match="w:numPr[not(following-sibling::w:numPr)]" mode="docx2hub:props2atts">
    <w:numPr>
      <xsl:apply-templates select="(w:numId, preceding-sibling::w:numPr/w:numId)[1]" mode="#current"/>
      <xsl:apply-templates select="(w:ilvl, preceding-sibling::w:numPr/w:ilvl)[1]" mode="#current"/>
      <xsl:apply-templates select="node() except w:numId, w:ilvl" mode="#current"/>
    </w:numPr>
  </xsl:template>
   
  <xsl:template match="docx2hub:remove-attribute" mode="docx2hub:props2atts" />

  <xsl:template match="docx2hub:style-link" mode="docx2hub:props2atts">
    <xsl:attribute name="{if (@type eq 'AppliedParagraphStyle')
                          then 'parastyle'
                          else @type}" 
      select="@target" />
  </xsl:template>

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- mode: docx2hub:remove-redundant-run-atts -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <!-- there are also templates of this mode in other files -->

  <xsl:template match="w:r/@*[some $att in ../ancestor::w:p[1]/@* 
                       satisfies (
                         name($att) eq name(current())
                         and 
                         xs:string($att) eq xs:string(current())
                       )]" mode="docx2hub:remove-redundant-run-atts" />

  <!-- commented in order to keep italics in index terms -->
  <!--<xsl:template match="w:r[*]
                          [every $c in * satisfies ($c/(self::w:instrText | self::w:fldChar))]
                          /@*[matches(name(), '^(css:|xml:lang)')]" mode="docx2hub:remove-redundant-run-atts" 
                priority="10"/>-->
  
  <!-- collateral: denote numbering resets -->
  <xsl:template match="w:p" mode="docx2hub:remove-redundant-run-atts">
    <xsl:param name="css:orientation" as="xs:string?" tunnel="yes"/>
    <xsl:copy>
      <xsl:variable name="style" select="key('docx2hub:style-by-role', @role, $root)[1]" as="element(css:rule)?"/>
      <xsl:variable name="numId" as="element(w:numId)?" 
        select="(w:numPr/w:numId, $style/w:numPr/w:numId)[1]"/>
      <xsl:variable name="ilvl" as="xs:integer" 
        select="(
                  for $i in 
                    (w:numPr/w:ilvl/@w:val,
                     $style/w:numPr/w:ilvl/@w:val)[1] 
                  return xs:integer($i),
                  0
                )[1]"/>
      <xsl:if test="$css:orientation">
        <xsl:attribute name="css:orientation" select="$css:orientation"/>
      </xsl:if>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates select="$numId" mode="docx2hub:abstractNum">
        <xsl:with-param name="ilvl" select="$ilvl"/>
      </xsl:apply-templates>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:tbl" mode="docx2hub:remove-redundant-run-atts">      
    <xsl:param name="css:orientation" as="xs:string?" tunnel="yes"/>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:if test="$css:orientation">
        <xsl:attribute name="css:orientation" select="$css:orientation"/>  
      </xsl:if>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>

<!--  collateral: remove w:proofErr (may cause problems in field functions)-->
  <xsl:template match="w:proofErr" mode="docx2hub:add-props"/>

  <xsl:template match="*[w:p[w:pgSz]]" mode="docx2hub:remove-redundant-run-atts">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="node()" group-ending-with="w:p[w:pgSz[@css:orientation='landscape']]">
        <xsl:choose>
          <xsl:when test="current-group()[last()][self::w:p[w:pgSz[@css:orientation='landscape']]]">
            <xsl:for-each-group select="current-group()" group-starting-with="w:p[w:pgSz[not(@css:orientation='landscape')]]">
              <xsl:choose>
                <xsl:when test="current-group()[last()][self::w:p[w:pgSz[@css:orientation='landscape']]]">
                  <xsl:apply-templates select="current-group()[1]" mode="#current"/>
                  <xsl:apply-templates select="current-group()[position() gt 1]" mode="docx2hub:remove-redundant-run-atts">
                    <xsl:with-param name="css:orientation" select="'landscape'" tunnel="yes"/>
                  </xsl:apply-templates>
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
  
  <xsl:template match="w:tbl/w:tr[every $i 
                                  in w:tc 
                                  satisfies (not(exists($i/@css:border-bottom-style)) and 
                                             not(exists($i/@css:border-bottom-width)) and 
                                             $i/w:tcPr[w:vMerge[not(exists(@w:val)) or (@w:val  ne 'restart')]] and 
                                             $i/w:p[not(child::node()[not(self::w:pPr)])])]" 
                mode="docx2hub:remove-redundant-run-atts">
  </xsl:template>
  
  <!-- preserve mml text nodes -->
  
  <xsl:template match="mml:*" mode="docx2hub:add-props" exclude-result-prefixes="#all">
    <xsl:copy inherit-namespaces="no" exclude-result-prefixes="#all">
      <xsl:apply-templates select="@*|node()" mode="#current"/>  
    </xsl:copy>
  </xsl:template>
  
</xsl:stylesheet>
